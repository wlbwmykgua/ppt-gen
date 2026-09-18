#!/usr/bin/env python3
"""Create, validate, and resume a ppt-gen project manifest."""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import shutil
import tempfile
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path
from typing import Any, Iterator

from artifact_contracts import CLI_DELIVERABLE_ROLES
from speech_timing import number as timing_number
from file_lock import exclusive_lock


SCHEMA_VERSION = 2
STAGES = [
    "brief",
    "report",
    "outline",
    "style-options",
    "template",
    "image-deck",
    "editable-deck",
    "deck-revision",
    "speaker-notes",
    "qa-package",
]
DELIVERABLE_STAGE = {
    "report-docx": "report",
    "outline": "outline",
    "style-options": "style-options",
    "template-overview": "template",
    "slide-images": "image-deck",
    "image-pptx": "image-deck",
    "editable-pptx": "editable-deck",
    "revised-pptx": "deck-revision",
    "speaker-notes-docx": "speaker-notes",
    "ppt-notes": "speaker-notes",
    "qa-report": "qa-package",
}
DELIVERABLE_ROLES = list(DELIVERABLE_STAGE)
REQUIRED_QA_ARTIFACT_ROLES = {
    "report-docx": {"report", "report-render-evidence"},
    "outline": {"outline", "slide-copy-ledger"},
    "style-options": {"style-options"},
    "template-overview": {"visual-system", "template-layouts", "template-overview", "template-composer"},
    "slide-images": {"slide-images", "slide-copy-ledger", "slide-manifest"},
    "image-pptx": {"image-pptx", "image-pptx-render-evidence", "slide-copy-ledger", "slide-manifest"},
    "editable-pptx": {"editable-pptx", "editable-validation", "editable-handoff", "offline-preparation", "slide-copy-ledger", "slide-manifest"},
    "revised-pptx": {"original-pptx", "revised-pptx", "revision-checkpoints"},
    "speaker-notes-docx": {
        "speaker-notes",
        "speaker-notes-render-evidence",
        "speaker-notes-ledger",
        "slide-copy-ledger",
        "reference-pptx",
    },
    "ppt-notes": {"ppt-notes-pptx", "speaker-notes-ledger", "slide-copy-ledger", "reference-pptx"},
}


def required_qa_roles(role: str, validation_contract_version: int = 2) -> set[str]:
    roles = REQUIRED_QA_ARTIFACT_ROLES.get(role, set()) | {"source-manifest", "claim-ledger"}
    # Preserve historical acceptance when loading existing projects. The new
    # validator always emits v2 and independently requires offline execution.
    if validation_contract_version == 1 and role == "editable-pptx":
        roles.discard("offline-preparation")
    return roles


STAGE_STATES = {
    "pending",
    "running",
    "waiting_user",
    "completed",
    "skipped",
    "stale",
    "failed",
    "cancelled",
}
PAGE_STATES = {
    "pending",
    "running",
    "waiting_user",
    "completed",
    "stale",
    "failed",
    "cancelled",
}
QA_STATES = {"not_run", "passed", "passed_with_warnings", "failed"}
PASSING_QA = {"passed", "passed_with_warnings"}
RUN_STATES = {"active", "paused", "completed", "blocked", "cancelled"}
MODES = {"guided", "stepwise", "continuous"}
SOURCE_POLICIES = {"faithful", "verify-update"}
DATA_HANDLING_POLICIES = {"standard", "local-only"}
RECORD_STATES = {"current", "stale", "drifted", "missing", "unsupported", "unrequested"}
MANIFEST_NAME = "ppt-gen-state.json"


def now() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


def manifest_path(project_dir: str | Path) -> Path:
    return Path(project_dir).expanduser().resolve() / MANIFEST_NAME


def parse_value(raw: str) -> Any:
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        return raw


def parse_pairs(items: list[str] | None, label: str) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for item in items or []:
        if "=" not in item:
            raise SystemExit(f"{label} must use KEY=VALUE: {item}")
        key, value = item.split("=", 1)
        key = key.strip()
        if not key:
            raise SystemExit(f"{label} key cannot be empty: {item}")
        result[key] = parse_value(value.strip())
    return result


def canonical_targets(raw_targets: list[str], start_at: str, stop_after: str) -> list[str]:
    start_idx = STAGES.index(start_at)
    stop_idx = STAGES.index(stop_after)
    if stop_idx < start_idx:
        raise SystemExit("--stop-after cannot precede --start-at")
    requested = set(raw_targets)
    outside = [
        stage
        for stage in requested
        if stage not in STAGES or not start_idx <= STAGES.index(stage) <= stop_idx
    ]
    if outside:
        raise SystemExit(f"--target stages fall outside start/stop range: {sorted(outside)}")
    targets = [stage for stage in STAGES if stage in requested]
    if not targets:
        raise SystemExit("At least one target stage is required")
    return targets


def default_route_targets(start_at: str, stop_after: str) -> list[str]:
    start_idx = STAGES.index(start_at)
    stop_idx = STAGES.index(stop_after)
    route = STAGES[start_idx : stop_idx + 1]
    if start_at != "deck-revision" and stop_after != "deck-revision":
        route = [stage for stage in route if stage != "deck-revision"]
    return route


def canonical_deliverables(raw_deliverables: list[str]) -> list[str]:
    unknown = sorted(set(raw_deliverables) - set(DELIVERABLE_ROLES))
    if unknown:
        raise SystemExit(f"Unknown deliverable roles: {unknown}")
    return [role for role in DELIVERABLE_ROLES if role in set(raw_deliverables)]


def validate_deliverable_route(deliverables: list[str], targets: list[str]) -> None:
    impossible = [role for role in deliverables if DELIVERABLE_STAGE[role] not in targets]
    if impossible:
        details = ", ".join(f"{role} requires {DELIVERABLE_STAGE[role]}" for role in impossible)
        raise SystemExit(f"Declared deliverables require missing target stages: {details}")


def _walk_error(error: OSError) -> None:
    raise error


def _update_file_digest(digest: Any, path: Path) -> int:
    size = 0
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
            size += len(chunk)
    return size


def file_record(raw_path: str) -> dict[str, Any]:
    path = Path(raw_path).expanduser().resolve()
    captured = now()
    record: dict[str, Any] = {
        "path": str(path),
        "exists": path.exists(),
        "captured_at": captured,
    }
    if not path.exists():
        record.update({"kind": "missing", "status": "missing"})
        return record

    try:
        if path.is_file():
            digest = hashlib.sha256()
            size = _update_file_digest(digest, path)
            stat = path.stat()
            record.update(
                {
                    "kind": "file",
                    "status": "current",
                    "sha256": digest.hexdigest(),
                    "size": size,
                    "file_count": 1,
                    "modified_at": datetime.fromtimestamp(stat.st_mtime)
                    .astimezone()
                    .isoformat(timespec="seconds"),
                }
            )
            return record

        if path.is_dir():
            digest = hashlib.sha256()
            digest.update(b"ppt-gen-directory-v1\0")
            total_size = 0
            file_count = 0
            latest_mtime = path.stat().st_mtime
            digest.update(b"D\0.\0")
            for current, dirnames, filenames in os.walk(path, followlinks=False, onerror=_walk_error):
                current_path = Path(current)
                dirnames.sort()
                filenames.sort()
                for dirname in list(dirnames):
                    child = current_path / dirname
                    relative = child.relative_to(path).as_posix()
                    latest_mtime = max(latest_mtime, child.lstat().st_mtime)
                    if child.is_symlink():
                        digest.update(f"L\0{relative}\0{os.readlink(child)}\0".encode("utf-8"))
                        dirnames.remove(dirname)
                    else:
                        digest.update(f"D\0{relative}\0".encode("utf-8"))
                for filename in filenames:
                    child = current_path / filename
                    relative = child.relative_to(path).as_posix()
                    stat = child.lstat()
                    latest_mtime = max(latest_mtime, stat.st_mtime)
                    if child.is_symlink():
                        digest.update(f"L\0{relative}\0{os.readlink(child)}\0".encode("utf-8"))
                        file_count += 1
                        continue
                    if not child.is_file():
                        digest.update(f"O\0{relative}\0{stat.st_mode}\0".encode("utf-8"))
                        continue
                    digest.update(f"F\0{relative}\0".encode("utf-8"))
                    total_size += _update_file_digest(digest, child)
                    file_count += 1
            record.update(
                {
                    "kind": "directory",
                    "status": "current",
                    "sha256": digest.hexdigest(),
                    "size": total_size,
                    "file_count": file_count,
                    "modified_at": datetime.fromtimestamp(latest_mtime)
                    .astimezone()
                    .isoformat(timespec="seconds"),
                }
            )
            return record

        record.update({"kind": "other", "status": "unsupported"})
        return record
    except OSError as exc:
        raise SystemExit(f"Cannot hash {path}: {exc}") from exc


def record_fingerprint(record: dict[str, Any]) -> tuple[Any, ...]:
    return (
        record.get("exists"),
        record.get("kind"),
        record.get("sha256"),
        record.get("size"),
        record.get("file_count"),
    )


def record_is_available(record: dict[str, Any]) -> bool:
    if not isinstance(record, dict):
        return False
    return (
        record.get("exists") is True
        and record.get("kind") in {"file", "directory"}
        and record.get("status", "current") == "current"
    )


def record_is_internal_authority(record: dict[str, Any]) -> bool:
    if not isinstance(record, dict):
        return False
    return (
        record.get("exists") is True
        and record.get("kind") in {"file", "directory"}
        and record.get("status", "current") in {"current", "unrequested"}
    )


def ensure_existing_record(raw_path: str, label: str) -> dict[str, Any]:
    record = file_record(raw_path)
    if not record_is_available(record):
        raise SystemExit(f"{label} must exist and be a regular file or directory: {record['path']}")
    return record


def ensure_not_self_referential(data: dict[str, Any], raw_path: str, label: str) -> None:
    candidate = Path(raw_path).expanduser().resolve()
    project = Path(data["project_dir"]).expanduser().resolve()
    state_path = project / MANIFEST_NAME
    lock_path = state_path.with_suffix(state_path.suffix + ".lock")
    if candidate in {state_path, lock_path} or (candidate.is_dir() and state_path.is_relative_to(candidate)):
        raise SystemExit(
            f"{label} cannot be the state file or a directory containing it: {candidate}"
        )


def expected_stage_dependencies(targets: list[str], stage: str) -> list[str]:
    if stage not in targets:
        return []
    index = targets.index(stage)
    preceding = targets[:index]
    if stage == "editable-deck":
        return ["image-deck"] if "image-deck" in preceding else []
    if stage == "speaker-notes":
        authority = next(
            (
                candidate
                for candidate in reversed(preceding)
                if candidate in {"deck-revision", "editable-deck", "image-deck"}
            ),
            None,
        )
        return [authority] if authority else []
    if stage == "qa-package":
        return list(preceding)
    if stage in {"brief", "deck-revision"} or not preceding:
        return []
    return [preceding[-1]]


def refresh_dependencies(data: dict[str, Any]) -> None:
    targets = data["targets"]
    for stage in STAGES:
        data["stages"][stage]["depends_on"] = expected_stage_dependencies(targets, stage)


def first_unresolved_target(data: dict[str, Any]) -> str | None:
    for stage in data["targets"]:
        if data["stages"][stage]["status"] != "completed":
            return stage
    return None


def selected_style(data: dict[str, Any]) -> Any:
    stages = data.get("stages") if isinstance(data.get("stages"), dict) else {}
    style_item = stages.get("style-options") if isinstance(stages.get("style-options"), dict) else {}
    preferences = data.get("preferences") if isinstance(data.get("preferences"), dict) else {}
    stage_value = style_item.get("selection")
    return stage_value if stage_value not in {None, ""} else preferences.get("selected_style")


def page_completion_errors(page_id: str, item: dict[str, Any]) -> list[str]:
    errors: list[str] = []
    if item.get("qa") not in PASSING_QA:
        errors.append(f"page {page_id} QA must be passed or passed_with_warnings")
    artifacts = item.get("artifacts", {})
    if not isinstance(artifacts, dict):
        errors.append(f"page {page_id} artifacts must be an object")
    else:
        for key, record in artifacts.items():
            if not record_is_available(record):
                errors.append(f"page {page_id} artifact {key!r} is missing, unsupported, or stale")
    return errors


def qa_report_errors(
    path_value: Any,
    label: str = "qa-report",
    state: dict[str, Any] | None = None,
    state_path: Path | None = None,
) -> list[str]:
    errors: list[str] = []
    if not isinstance(path_value, str):
        return [f"{label} path is missing"]
    path = Path(path_value).expanduser().resolve()
    if not path.is_file():
        return [f"{label} file is missing: {path}"]
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as exc:
        return [f"{label} is not valid JSON: {exc}"]
    if not isinstance(payload, dict):
        return [f"{label} must be a JSON object"]
    if payload.get("schema_version") != 1:
        errors.append(f"{label} schema_version must be 1")
    contract_version = payload.get("validation_contract_version", 1)
    if not isinstance(contract_version, int) or isinstance(contract_version, bool) or contract_version not in {1, 2}:
        errors.append(f"{label} has an unsupported validation_contract_version")
    if payload.get("tool") != "ppt-gen.validate_delivery":
        errors.append(f"{label} tool must be ppt-gen.validate_delivery")
    if payload.get("mode") != "final-delivery":
        errors.append(f"{label} mode must be final-delivery; structural-only evidence is diagnostic")
    if payload.get("passed") is not True or payload.get("status") not in PASSING_QA:
        errors.append(f"{label} must have a passing final status")
    if payload.get("errors") not in ([], None):
        errors.append(f"{label} contains hard errors")
    if not isinstance(payload.get("checked_at"), str) or not payload.get("checked_at"):
        errors.append(f"{label} checked_at must be a non-empty timestamp")
    gates = payload.get("gates")
    if not isinstance(gates, list) or not gates:
        errors.append(f"{label} must contain nonempty gates")
    else:
        for index, gate in enumerate(gates, 1):
            if not isinstance(gate, dict) or not isinstance(gate.get("id"), str):
                errors.append(f"{label} gate {index} is malformed")
                continue
            if gate.get("status") not in PASSING_QA:
                errors.append(f"{label} gate {gate.get('id')!r} is not passing")
            if not isinstance(gate.get("evidence"), list) or not gate.get("evidence"):
                errors.append(f"{label} gate {gate.get('id')!r} has no evidence")
    qa_warnings = payload.get("warnings")
    if not isinstance(qa_warnings, list):
        errors.append(f"{label} warnings must be an array")
        qa_warnings = []
    normalized_warnings = [str(value).strip() for value in qa_warnings if str(value).strip()]
    if payload.get("status") == "passed_with_warnings" and not normalized_warnings:
        errors.append(f"{label} status is passed_with_warnings but warnings are empty")
    if normalized_warnings and payload.get("status") != "passed_with_warnings":
        errors.append(f"{label} contains warnings but status is not passed_with_warnings")
    authorities = payload.get("authorities")
    if not isinstance(authorities, dict) or not {"source-manifest", "claim-ledger"}.issubset(authorities):
        errors.append(f"{label} must bind source-manifest and claim-ledger authorities")
    artifacts = payload.get("artifacts")
    if not isinstance(artifacts, list) or not artifacts:
        errors.append(f"{label} must contain hashed artifacts")
    if state is not None and state_path is not None:
        checkpoint = payload.get("state_checkpoint")
        if not isinstance(checkpoint, dict):
            errors.append(f"{label} must contain state_checkpoint")
        else:
            checkpoint_path = checkpoint.get("path")
            if not isinstance(checkpoint_path, str) or Path(checkpoint_path).expanduser().resolve() != state_path.resolve():
                errors.append(f"{label} state_checkpoint.path does not match this state")
            history = state.get("history") if isinstance(state.get("history"), list) else []
            verify_id = checkpoint.get("verify_history_id")
            matched_verify = next(
                (
                    entry for entry in history
                    if isinstance(entry, dict)
                    and entry.get("event") == "verified"
                    and entry.get("id") == verify_id
                ),
                None,
            )
            if matched_verify is None:
                errors.append(f"{label} verify checkpoint does not identify a verify event in this state")
            else:
                verify_details = matched_verify.get("details") if isinstance(matched_verify.get("details"), dict) else {}
                if checkpoint.get("verify_revision") != matched_verify.get("revision"):
                    errors.append(f"{label} verify revision does not match its verify event")
                if checkpoint.get("state_revision") != matched_verify.get("revision"):
                    errors.append(f"{label} state revision does not match its verified revision")
                if checkpoint.get("drift_count") != 0 or verify_details.get("drift_count") != 0:
                    errors.append(f"{label} state checkpoint contains drift")
        state_records: list[dict[str, Any]] = []
        for key in ("inputs", "artifacts", "deliverable_artifacts"):
            container = state.get(key)
            if isinstance(container, dict):
                state_records.extend(record for record in container.values() if isinstance(record, dict))
        stages = state.get("stages")
        if isinstance(stages, dict):
            for stage in stages.values():
                if not isinstance(stage, dict):
                    continue
                stage_artifacts = stage.get("artifacts")
                if isinstance(stage_artifacts, dict):
                    state_records.extend(record for record in stage_artifacts.values() if isinstance(record, dict))
        for authority_name in ("source-manifest", "claim-ledger"):
            authority = authorities.get(authority_name) if isinstance(authorities, dict) else None
            if not isinstance(authority, dict):
                continue
            authority_path = authority.get("path")
            authority_hash = authority.get("sha256")
            matched = any(
                isinstance(record, dict)
                and isinstance(record.get("path"), str)
                and isinstance(record.get("sha256"), str)
                and isinstance(authority_path, str)
                and Path(record["path"]).expanduser().resolve() == Path(authority_path).expanduser().resolve()
                and record["sha256"].casefold() == str(authority_hash).casefold()
                for record in state_records
            )
            if not matched:
                errors.append(f"{label} authority {authority_name!r} is not current in this state")
            if (
                isinstance(authority_path, str)
                and isinstance(authority_hash, str)
            ):
                authority_file = Path(authority_path).expanduser().resolve()
                try:
                    current_authority = file_record(authority_file)
                except OSError as exc:
                    errors.append(f"{label} authority {authority_name!r} cannot be rehashed: {exc}")
                else:
                    if (
                        not record_is_available(current_authority)
                        or current_authority.get("sha256") != authority_hash.casefold()
                    ):
                        errors.append(f"{label} authority {authority_name!r} is missing or its current SHA-256 mismatches")
        if isinstance(artifacts, list):
            for index, artifact in enumerate(artifacts, 1):
                if not isinstance(artifact, dict):
                    errors.append(f"{label} artifact {index} must be an object")
                    continue
                role = artifact.get("role")
                artifact_path_value = artifact.get("path")
                artifact_hash = artifact.get("sha256")
                if not isinstance(role, str) or not role.strip():
                    errors.append(f"{label} artifact {index} has no role")
                if not isinstance(artifact_path_value, str) or not isinstance(artifact_hash, str):
                    errors.append(f"{label} artifact {index} must bind path and SHA-256")
                    continue
                artifact_path = Path(artifact_path_value).expanduser().resolve()
                try:
                    current = file_record(artifact_path)
                except OSError as exc:
                    errors.append(f"{label} artifact {index} cannot be rehashed: {exc}")
                    continue
                if not record_is_available(current) or current.get("sha256") != artifact_hash.casefold():
                    errors.append(f"{label} artifact {index} is missing or its current SHA-256 mismatches")
                    continue
                registered = any(
                    isinstance(record.get("path"), str)
                    and Path(record["path"]).expanduser().resolve() == artifact_path
                    and isinstance(record.get("sha256"), str)
                    and record["sha256"].casefold() == artifact_hash.casefold()
                    and record_is_available(record)
                    for record in state_records
                )
                if not registered:
                    errors.append(f"{label} artifact {role!r} is not registered current in this state")
        source_authority = authorities.get("source-manifest") if isinstance(authorities, dict) else None
        if isinstance(source_authority, dict) and isinstance(source_authority.get("path"), str):
            source_path = Path(source_authority["path"]).expanduser().resolve()
            try:
                source_payload = json.loads(source_path.read_text(encoding="utf-8"))
            except (OSError, UnicodeError, json.JSONDecodeError):
                source_payload = None
            policies = state.get("policies") if isinstance(state.get("policies"), dict) else {}
            if isinstance(source_payload, dict):
                for key in ("source_policy", "data_handling"):
                    if source_payload.get(key) != policies.get(key):
                        errors.append(f"{label} source-manifest {key} conflicts with state policies")
    return errors


def qa_report_artifact_bindings(path_value: Any) -> dict[Path, str]:
    if not isinstance(path_value, str):
        return {}
    path = Path(path_value).expanduser().resolve()
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError):
        return {}
    bindings: dict[Path, str] = {}
    for artifact in payload.get("artifacts", []) if isinstance(payload, dict) else []:
        if not isinstance(artifact, dict):
            continue
        artifact_path = artifact.get("path")
        digest = artifact.get("sha256")
        if isinstance(artifact_path, str) and isinstance(digest, str):
            bindings[Path(artifact_path).expanduser().resolve()] = digest.casefold()
    return bindings


def stage_completion_errors(data: dict[str, Any], stage: str) -> list[str]:
    item = data["stages"][stage]
    errors: list[str] = []
    policies = data.get("policies") if isinstance(data.get("policies"), dict) else {}
    if policies.get("source_policy") not in SOURCE_POLICIES:
        errors.append("source_policy must be recorded as faithful or verify-update")
    if policies.get("data_handling") not in DATA_HANDLING_POLICIES:
        errors.append("data_handling must be recorded as standard or local-only")
    if policies.get("ocr_policy") != "offline" or policies.get("ask_for_ocr_token") is not False:
        errors.append("OCR policy must remain offline/no-token")
    if item.get("qa") not in PASSING_QA:
        errors.append("QA must be passed or passed_with_warnings")
    stage_warnings = item.get("warnings") if isinstance(item.get("warnings"), list) else []
    if stage_warnings and item.get("qa") != "passed_with_warnings":
        errors.append("a stage with recorded warnings must use QA status passed_with_warnings")
    artifacts = item.get("artifacts", {})
    if not isinstance(artifacts, dict):
        errors.append("artifacts must be an object")
        artifacts = {}
    else:
        for key, record in artifacts.items():
            if not record_is_available(record):
                errors.append(f"artifact {key!r} is missing, unsupported, or stale")
    declared_roles = [
        role for role in data.get("deliverables", []) if DELIVERABLE_STAGE.get(role) == stage
    ]
    deliverable_artifacts = data.get("deliverable_artifacts", {})
    if not isinstance(deliverable_artifacts, dict):
        deliverable_artifacts = {}
    for role in declared_roles:
        record = deliverable_artifacts.get(role)
        if record is None:
            errors.append(f"declared deliverable {role!r} has no registered artifact")
        elif not record_is_available(record):
            errors.append(f"declared deliverable {role!r} is missing, unsupported, or stale")
    qa_records: list[tuple[str, dict[str, Any]]] = []
    for key, record in artifacts.items():
        if isinstance(record, dict) and ("qa" in key.casefold() or Path(str(record.get("path", ""))).name == "qa-report.json"):
            qa_records.append((f"artifact {key!r}", record))
    qa_deliverable = deliverable_artifacts.get("qa-report")
    if (
        stage == "qa-package"
        and isinstance(qa_deliverable, dict)
        and not any(
            Path(str(record.get("path", ""))).expanduser().resolve()
            == Path(str(qa_deliverable.get("path", ""))).expanduser().resolve()
            for _, record in qa_records
        )
    ):
        qa_records.append(("deliverable 'qa-report'", qa_deliverable))
    if qa_records:
        qa_has_warnings = False
        for qa_label, qa_record in qa_records:
            errors.extend(qa_report_errors(qa_record.get("path"), qa_label, data, Path(data["project_dir"]) / MANIFEST_NAME))
            try:
                qa_payload = json.loads(Path(str(qa_record.get("path"))).expanduser().resolve().read_text(encoding="utf-8"))
            except (OSError, UnicodeError, json.JSONDecodeError, TypeError):
                qa_payload = None
            if isinstance(qa_payload, dict):
                report_warnings = qa_payload.get("warnings")
                if qa_payload.get("status") == "passed_with_warnings" or (
                    isinstance(report_warnings, list)
                    and any(str(value).strip() for value in report_warnings)
                ):
                    qa_has_warnings = True
        expected_qa = "passed_with_warnings" if qa_has_warnings else "passed"
        if item.get("qa") != expected_qa:
            errors.append(
                f"stage QA status {item.get('qa')!r} does not match final qa-report status {expected_qa!r}"
            )
        try:
            accepted_payload = json.loads(
                Path(str(qa_records[-1][1].get("path"))).expanduser().resolve().read_text(encoding="utf-8")
            )
        except (OSError, UnicodeError, json.JSONDecodeError, TypeError):
            accepted_payload = None
        accepted_roles = {
            artifact.get("role")
            for artifact in accepted_payload.get("artifacts", [])
            if isinstance(accepted_payload, dict)
            and isinstance(artifact, dict)
            and isinstance(artifact.get("role"), str)
        } if isinstance(accepted_payload, dict) else set()
        accepted_roles.update({"source-manifest", "claim-ledger"})
        imported_roles = [
            CLI_DELIVERABLE_ROLES[record["artifact_role"]][0]
            for record in data.get("inputs", {}).values()
            if stage == "qa-package" and isinstance(record, dict)
            and record.get("authority_stage") == "qa-package"
            and isinstance(record.get("artifact_role"), str)
            and record.get("artifact_role") in CLI_DELIVERABLE_ROLES
        ]
        for role in [*declared_roles, *imported_roles]:
            contract_version = accepted_payload.get("validation_contract_version", 1) if isinstance(accepted_payload, dict) else 2
            required = required_qa_roles(role, contract_version)
            missing = sorted(required - accepted_roles)
            if missing:
                errors.append(f"final qa-report for {role!r} is missing required artifact roles: {missing}")
        qa_bindings = qa_report_artifact_bindings(qa_records[-1][1].get("path"))
        if stage == "qa-package":
            for key, record in data.get("inputs", {}).items():
                if not isinstance(record, dict) or not isinstance(record.get("artifact_role"), str) or record.get("artifact_role") not in CLI_DELIVERABLE_ROLES:
                    continue
                if qa_bindings.get(Path(record["path"]).resolve()) != record.get("sha256"):
                    errors.append(f"final qa-report does not bind imported QA input {key!r} by current path/hash")
        roles_to_bind = (
            [role for role in data.get("deliverables", []) if role != "qa-report"]
            if stage == "qa-package"
            else declared_roles
        )
        for role in roles_to_bind:
            if role == "qa-report":
                # The QA report is this very file; requiring its own digest in
                # its artifact list would create an impossible self-hash.
                continue
            record = deliverable_artifacts.get(role)
            if not isinstance(record, dict) or not isinstance(record.get("path"), str):
                continue
            artifact_path = Path(record["path"]).expanduser().resolve()
            expected_hash = record.get("sha256")
            if qa_bindings.get(artifact_path) != (expected_hash.casefold() if isinstance(expected_hash, str) else None):
                errors.append(f"final qa-report does not bind declared deliverable {role!r} by current path/hash")
    elif (declared_roles or stage == "qa-package") and data.get("qa_contract_version") == 1:
        errors.append("a final qa-report.json artifact is required to complete a deliverable stage")
    available_stage_artifacts = any(record_is_available(record) for record in artifacts.values())
    available_deliverables = any(
        record.get("stage") == stage and record_is_internal_authority(record)
        for record in deliverable_artifacts.values()
    )
    if not available_stage_artifacts and not available_deliverables:
        errors.append("at least one existing stage artifact or declared deliverable artifact is required")
    inputs = data.get("inputs", {})
    if isinstance(inputs, dict):
        for role, record in inputs.items():
            if isinstance(record, dict) and record.get("authority_stage") == stage:
                if not record_is_available(record):
                    errors.append(f"authority input {role!r} is missing, unsupported, or stale")
    pages = item.get("pages", {})
    if not isinstance(pages, dict):
        errors.append("pages must be an object")
    else:
        for page_id, page in pages.items():
            if not isinstance(page, dict):
                errors.append(f"page {page_id} must be an object")
            elif page.get("status") != "completed":
                errors.append(f"page {page_id} is {page.get('status')}, not completed")
            else:
                errors.extend(page_completion_errors(page_id, page))
    dependencies = item.get("depends_on", [])
    if not isinstance(dependencies, list):
        errors.append("depends_on must be a list")
        dependencies = []
    stages = data.get("stages", {})
    for dependency in dependencies:
        dependency_item = stages.get(dependency) if isinstance(stages, dict) else None
        if not isinstance(dependency_item, dict) or dependency_item.get("status") != "completed":
            errors.append(f"dependency {dependency!r} is not completed")
    if stage == "style-options" and not selected_style(data):
        downstream_visual = {"template", "image-deck", "editable-deck"}
        if downstream_visual.intersection(data.get("targets", [])):
            errors.append("a visual style must be selected before downstream visual targets")
    return errors


def derive_run_state(data: dict[str, Any], preferred: str | None = None) -> None:
    current = first_unresolved_target(data)
    data["current_stage"] = current
    if current is None:
        data["run_status"] = "completed"
        return
    current_status = data["stages"][current]["status"]
    if preferred == "completed":
        raise SystemExit("Cannot mark run completed while target stages remain unfinished")
    if preferred == "cancelled" or current_status == "cancelled":
        data["run_status"] = "cancelled"
    elif current_status == "waiting_user":
        data["run_status"] = "paused"
    elif current_status == "failed":
        data["run_status"] = "blocked"
    elif preferred in {"paused", "blocked"}:
        data["run_status"] = preferred
    else:
        data["run_status"] = "active"


def rebuild_artifact_index(data: dict[str, Any]) -> None:
    index: dict[str, Any] = {}
    for stage in STAGES:
        for key, record in data["stages"][stage].get("artifacts", {}).items():
            index[f"{stage}:{key}"] = {**record, "stage": stage}
    data["artifacts"] = index


def append_history(
    data: dict[str, Any], event: str, *, stage: str | None = None, details: dict[str, Any] | None = None
) -> None:
    history = data.setdefault("history", [])
    entry: dict[str, Any] = {
        "id": len(history) + 1,
        "at": now(),
        "revision": int(data.get("state_revision", 0)) + 1,
        "event": event,
    }
    if stage:
        entry["stage"] = stage
    if details:
        entry["details"] = details
    history.append(entry)


def mark_stage_stale(data: dict[str, Any], stage: str, reason: str | None = None) -> bool:
    if stage not in data["targets"]:
        return False
    item = data["stages"][stage]
    if item["status"] in {"skipped", "cancelled"}:
        return False
    changed = item["status"] != "stale" or item["qa"] != "not_run"
    item["status"] = "stale"
    item["qa"] = "not_run"
    item["revision"] = int(item.get("revision", 0)) + 1
    if reason and reason not in item["warnings"]:
        item["warnings"].append(reason)
    for record in item.get("artifacts", {}).values():
        if record.get("status") == "current":
            record["status"] = "stale"
    for page in item.get("pages", {}).values():
        if page.get("status") not in {"cancelled"}:
            page["status"] = "stale"
            page["qa"] = "not_run"
            page["revision"] = int(page.get("revision", 0)) + 1
        for record in page.get("artifacts", {}).values():
            if record.get("status") == "current":
                record["status"] = "stale"
    for deliverable in data.get("deliverable_artifacts", {}).values():
        if deliverable.get("stage") == stage and deliverable.get("status") == "current":
            deliverable["status"] = "stale"
    return changed


def invalidate_from(
    data: dict[str, Any], stage: str, *, include_stage: bool, reason: str | None = None
) -> list[str]:
    affected: list[str] = []
    consumers: set[str] = set()
    frontier = [stage]
    while frontier:
        authority = frontier.pop(0)
        for candidate in data["targets"]:
            if candidate in consumers or candidate == stage:
                continue
            if authority in data["stages"][candidate].get("depends_on", []):
                consumers.add(candidate)
                frontier.append(candidate)
    for candidate in data["targets"]:
        if (include_stage and candidate == stage) or candidate in consumers:
            if mark_stage_stale(data, candidate, reason):
                affected.append(candidate)
    rebuild_artifact_index(data)
    return affected


def validate_record(record: Any, label: str, errors: list[str]) -> None:
    if not isinstance(record, dict):
        errors.append(f"{label} must be an object")
        return
    if not isinstance(record.get("path"), str):
        errors.append(f"{label}.path must be a string")
    if not isinstance(record.get("exists"), bool):
        errors.append(f"{label}.exists must be a boolean")
    if record.get("status") not in RECORD_STATES:
        errors.append(f"{label}.status is invalid")


def manifest_errors(data: Any) -> list[str]:
    errors: list[str] = []
    if not isinstance(data, dict):
        return ["manifest root must be an object"]
    if data.get("schema_version") != SCHEMA_VERSION:
        errors.append(f"schema_version must be {SCHEMA_VERSION}")
    if data.get("qa_contract_version") != 1:
        errors.append("qa_contract_version must be 1")
    if not isinstance(data.get("state_revision"), int) or data.get("state_revision", -1) < 0:
        errors.append("state_revision must be a non-negative integer")
    if data.get("run_mode") not in MODES:
        errors.append("run_mode is invalid")
    if data.get("run_status") not in RUN_STATES:
        errors.append("run_status is invalid")
    if not isinstance(data.get("project_name"), str) or not data.get("project_name"):
        errors.append("project_name must be a non-empty string")
    if not isinstance(data.get("project_dir"), str) or not data.get("project_dir"):
        errors.append("project_dir must be a non-empty string")
    if not isinstance(data.get("preferences"), dict):
        errors.append("preferences must be an object")
    elif data["preferences"].get("offline_ocr") is not True:
        errors.append("preferences.offline_ocr must be true")
    policies = data.get("policies")
    if policies is not None:
        if not isinstance(policies, dict):
            errors.append("policies must be an object")
        else:
            source_policy = policies.get("source_policy")
            if source_policy is not None and source_policy not in SOURCE_POLICIES:
                errors.append("policies.source_policy is invalid")
            if policies.get("data_handling") not in DATA_HANDLING_POLICIES:
                errors.append("policies.data_handling is invalid")
            if policies.get("ocr_policy") != "offline":
                errors.append("policies.ocr_policy must be offline")
            if policies.get("ask_for_ocr_token") is not False:
                errors.append("policies.ask_for_ocr_token must be false")
    for field in ("assumptions", "decisions", "history"):
        if not isinstance(data.get(field), list):
            errors.append(f"{field} must be a list")
    for field in ("created_at", "updated_at"):
        if not isinstance(data.get(field), str) or not data.get(field):
            errors.append(f"{field} must be a non-empty timestamp string")
    if data.get("start_at") not in STAGES or data.get("stop_after") not in STAGES:
        errors.append("start_at and stop_after must be valid stages")
    elif STAGES.index(data["stop_after"]) < STAGES.index(data["start_at"]):
        errors.append("stop_after precedes start_at")

    raw_targets = data.get("targets")
    if not isinstance(raw_targets, list) or not raw_targets:
        errors.append("targets must be a non-empty list")
        targets = []
    else:
        targets = [stage for stage in raw_targets if isinstance(stage, str)]
        if len(targets) != len(raw_targets):
            errors.append("every target must be a stage-name string")
        elif targets != [stage for stage in STAGES if stage in set(targets)]:
            errors.append("targets must be unique and in canonical stage order")
    for stage in targets:
        if stage not in STAGES:
            errors.append(f"unknown target stage {stage!r}")

    stages = data.get("stages")
    if not isinstance(stages, dict):
        errors.append("stages must be an object")
        stages = {}
    for stage in STAGES:
        item = stages.get(stage)
        if not isinstance(item, dict):
            errors.append(f"stage {stage!r} is missing or invalid")
            continue
        if item.get("status") not in STAGE_STATES:
            errors.append(f"stage {stage!r} has an invalid status")
        if item.get("qa") not in QA_STATES:
            errors.append(f"stage {stage!r} has an invalid QA state")
        if not isinstance(item.get("revision"), int):
            errors.append(f"stage {stage!r}.revision must be an integer")
        for field in ("warnings", "notes"):
            if not isinstance(item.get(field), list):
                errors.append(f"stage {stage!r}.{field} must be a list")
        if not isinstance(item.get("depends_on"), list):
            errors.append(f"stage {stage!r}.depends_on must be a list")
        elif item.get("depends_on") != expected_stage_dependencies(targets, stage):
            errors.append(
                f"stage {stage!r}.depends_on must be "
                f"{expected_stage_dependencies(targets, stage)!r} for this route"
            )
        if stage in targets and item.get("status") == "skipped":
            errors.append(f"target stage {stage!r} cannot be skipped")
        if stage not in targets and item.get("status") != "skipped":
            errors.append(f"non-target stage {stage!r} must be skipped")
        artifacts = item.get("artifacts")
        if not isinstance(artifacts, dict):
            errors.append(f"stage {stage!r}.artifacts must be an object")
        else:
            for key, record in artifacts.items():
                validate_record(record, f"stages.{stage}.artifacts.{key}", errors)
        pages = item.get("pages")
        if not isinstance(pages, dict):
            errors.append(f"stage {stage!r}.pages must be an object")
        else:
            for page_id, page in pages.items():
                if not isinstance(page, dict):
                    errors.append(f"stages.{stage}.pages.{page_id} must be an object")
                    continue
                if page.get("status") not in PAGE_STATES:
                    errors.append(f"stages.{stage}.pages.{page_id}.status is invalid")
                if page.get("qa") not in QA_STATES:
                    errors.append(f"stages.{stage}.pages.{page_id}.qa is invalid")
                if not isinstance(page.get("revision"), int):
                    errors.append(f"stages.{stage}.pages.{page_id}.revision must be an integer")
                for field in ("warnings", "notes"):
                    if not isinstance(page.get(field), list):
                        errors.append(f"stages.{stage}.pages.{page_id}.{field} must be a list")
                page_artifacts = page.get("artifacts")
                if not isinstance(page_artifacts, dict):
                    errors.append(f"stages.{stage}.pages.{page_id}.artifacts must be an object")
                else:
                    for key, record in page_artifacts.items():
                        validate_record(record, f"stages.{stage}.pages.{page_id}.artifacts.{key}", errors)
                if page.get("status") == "completed":
                    errors.extend(f"stage {stage!r}: {message}" for message in page_completion_errors(page_id, page))
        if item.get("status") == "completed":
            errors.extend(f"stage {stage!r}: {message}" for message in stage_completion_errors(data, stage))

    raw_deliverables = data.get("deliverables")
    if not isinstance(raw_deliverables, list):
        errors.append("deliverables must be a list of requested roles")
        deliverables = []
    else:
        deliverables = [role for role in raw_deliverables if isinstance(role, str)]
        if len(deliverables) != len(raw_deliverables):
            errors.append("every deliverable must be a role-name string")
        unknown_deliverables = sorted(set(deliverables) - set(DELIVERABLE_ROLES))
        if unknown_deliverables:
            errors.append(f"unknown deliverable roles: {unknown_deliverables}")
        elif deliverables != [role for role in DELIVERABLE_ROLES if role in set(deliverables)]:
            errors.append("deliverables must be unique and in canonical role order")
        for role in deliverables:
            if role in DELIVERABLE_STAGE and DELIVERABLE_STAGE[role] not in targets:
                errors.append(f"deliverable {role!r} requires target {DELIVERABLE_STAGE[role]!r}")

    deliverable_artifacts = data.get("deliverable_artifacts")
    if not isinstance(deliverable_artifacts, dict):
        errors.append("deliverable_artifacts must be an object")
    else:
        for role, record in deliverable_artifacts.items():
            validate_record(record, f"deliverable_artifacts.{role}", errors)
            if role not in DELIVERABLE_STAGE:
                errors.append(f"deliverable_artifacts contains unknown role {role!r}")
                continue
            if isinstance(record, dict) and record.get("stage") != DELIVERABLE_STAGE[role]:
                errors.append(
                    f"deliverable_artifacts.{role}.stage must be {DELIVERABLE_STAGE[role]!r}"
                )
            if role not in deliverables and isinstance(record, dict) and record.get("status") == "current":
                errors.append(f"unrequested deliverable artifact {role!r} cannot be current")

    inputs = data.get("inputs")
    if not isinstance(inputs, dict):
        errors.append("inputs must be an object")
    else:
        for role, record in inputs.items():
            validate_record(record, f"inputs.{role}", errors)
            if isinstance(record, dict) and record.get("authority_stage") not in targets:
                errors.append(f"inputs.{role}.authority_stage must name a target stage")
            if isinstance(record, dict) and "artifact_role" in record:
                if not isinstance(record["artifact_role"], str) or record["artifact_role"] not in CLI_DELIVERABLE_ROLES:
                    errors.append(f"inputs.{role}.artifact_role is not a supported QA artifact role")
                if record.get("authority_stage") != "qa-package":
                    errors.append(f"inputs.{role}.artifact_role requires authority_stage=qa-package")

    artifact_index = data.get("artifacts")
    if not isinstance(artifact_index, dict):
        errors.append("artifacts must be an object")
    else:
        for key, record in artifact_index.items():
            validate_record(record, f"artifacts.{key}", errors)
            if isinstance(record, dict) and record.get("stage") not in STAGES:
                errors.append(f"artifacts.{key}.stage is invalid")
        expected_index: dict[str, Any] = {}
        for stage in STAGES:
            item = stages.get(stage) if isinstance(stages, dict) else None
            stage_artifacts = item.get("artifacts") if isinstance(item, dict) else None
            if isinstance(stage_artifacts, dict):
                for key, record in stage_artifacts.items():
                    if isinstance(record, dict):
                        expected_index[f"{stage}:{key}"] = {**record, "stage": stage}
        if artifact_index != expected_index:
            errors.append("artifacts index does not match stage artifacts")

    if targets and stages:
        expected_current = None
        for stage in targets:
            if isinstance(stages.get(stage), dict) and stages[stage].get("status") != "completed":
                expected_current = stage
                break
        if data.get("current_stage") != expected_current:
            errors.append(f"current_stage must be derived as {expected_current!r}")
        if data.get("current_stage") is not None:
            current_item = stages.get(data["current_stage"], {})
            if current_item.get("status") == "skipped":
                errors.append("current_stage cannot point to a skipped stage")
        if data.get("run_status") == "completed":
            unfinished = [stage for stage in targets if stages.get(stage, {}).get("status") != "completed"]
            if unfinished:
                errors.append(f"completed run has unfinished targets: {unfinished}")
            if data.get("current_stage") is not None:
                errors.append("completed run must have current_stage=null")
        elif expected_current is None:
            errors.append("all targets are completed but run_status is not completed")
    return errors


def validate_manifest(data: Any, path: Path | None = None) -> None:
    errors = manifest_errors(data)
    if errors:
        where = f" {path}" if path else ""
        formatted = "\n".join(f"- {error}" for error in errors)
        raise SystemExit(f"Invalid state manifest{where}:\n{formatted}")


def backup_manifest(path: Path, reason: str) -> Path:
    stamp = datetime.now().astimezone().strftime("%Y%m%dT%H%M%S%z")
    candidate = path.with_name(f"{path.stem}.{reason}-backup-{stamp}{path.suffix}")
    counter = 1
    while candidate.exists():
        candidate = path.with_name(f"{path.stem}.{reason}-backup-{stamp}-{counter}{path.suffix}")
        counter += 1
    shutil.copy2(path, candidate)
    return candidate


def atomic_save(path: Path, data: dict[str, Any]) -> None:
    old_revision = int(data.get("state_revision", 0))
    data["state_revision"] = old_revision + 1
    data["updated_at"] = now()
    try:
        validate_manifest(data, path)
    except BaseException:
        data["state_revision"] = old_revision
        raise
    descriptor, raw_temp = tempfile.mkstemp(prefix=f".{path.name}.", suffix=".tmp", dir=path.parent)
    temp = Path(raw_temp)
    try:
        with os.fdopen(descriptor, "w", encoding="utf-8") as handle:
            json.dump(data, handle, ensure_ascii=False, indent=2)
            handle.write("\n")
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(temp, path)
        try:
            directory_fd = os.open(path.parent, os.O_RDONLY)
            try:
                os.fsync(directory_fd)
            finally:
                os.close(directory_fd)
        except OSError:
            pass
    except BaseException:
        if temp.exists():
            temp.unlink()
        data["state_revision"] = old_revision
        raise


@contextmanager
def state_lock(project_dir: str | Path, *, create: bool = False) -> Iterator[Path]:
    path = manifest_path(project_dir)
    if create:
        path.parent.mkdir(parents=True, exist_ok=True)
    if not path.parent.exists():
        raise SystemExit(f"Project directory not found: {path.parent}")
    lock_path = path.with_suffix(path.suffix + ".lock")
    with exclusive_lock(lock_path):
        yield path


def normalize_legacy_record(raw: Any) -> dict[str, Any]:
    if not isinstance(raw, dict) or not isinstance(raw.get("path"), str):
        return {"path": "", "exists": False, "kind": "missing", "status": "missing", "captured_at": now()}
    snapshot = file_record(raw["path"])
    old_sha = raw.get("sha256")
    if old_sha and snapshot.get("sha256") != old_sha:
        baseline = dict(raw)
        baseline.setdefault("kind", "file")
        baseline.setdefault("captured_at", raw.get("modified_at", now()))
        baseline["status"] = "missing" if not snapshot.get("exists") else "drifted"
        baseline["observed"] = snapshot
        return baseline
    return snapshot


def migrate_v1(raw: dict[str, Any], path: Path, backup: Path) -> dict[str, Any]:
    legacy_targets = raw.get("targets", []) if isinstance(raw.get("targets"), list) else []
    raw_targets = [stage for stage in legacy_targets if stage in STAGES and stage != "deck-revision"]
    start_at = raw.get("start_at") if raw.get("start_at") in STAGES else (raw_targets[0] if raw_targets else "brief")
    stop_after = raw.get("stop_after") if raw.get("stop_after") in STAGES else (raw_targets[-1] if raw_targets else "qa-package")
    if STAGES.index(stop_after) < STAGES.index(start_at):
        stop_after = start_at
    if not raw_targets:
        raw_targets = default_route_targets(start_at, stop_after)
    targets = canonical_targets(raw_targets, start_at, stop_after)
    old_stages = raw.get("stages", {}) if isinstance(raw.get("stages"), dict) else {}
    stages: dict[str, Any] = {}
    for stage in STAGES:
        old = old_stages.get(stage, {}) if isinstance(old_stages.get(stage), dict) else {}
        status = old.get("status", "pending" if stage in targets else "skipped")
        if stage not in targets:
            status = "skipped"
        elif status not in STAGE_STATES or status == "skipped":
            status = "pending"
        qa = old.get("qa", "not_run") if old.get("qa") in QA_STATES else "not_run"
        old_artifacts = old.get("artifacts", {}) if isinstance(old.get("artifacts"), dict) else {}
        artifacts = {
            key: normalize_legacy_record(record)
            for key, record in old_artifacts.items()
            if isinstance(key, str)
        }
        stages[stage] = {
            "status": status,
            "qa": qa,
            "artifacts": artifacts,
            "pages": {},
            "warnings": list(old.get("warnings", [])) if isinstance(old.get("warnings"), list) else [],
            "notes": list(old.get("notes", [])) if isinstance(old.get("notes"), list) else [],
            "depends_on": [],
            "selection": None,
            "revision": 0,
        }

    preferences = raw.get("preferences", {}) if isinstance(raw.get("preferences"), dict) else {}
    data: dict[str, Any] = {
        "schema_version": SCHEMA_VERSION,
        "qa_contract_version": 1,
        "state_revision": raw.get("state_revision", 0)
        if isinstance(raw.get("state_revision", 0), int)
        else 0,
        "project_name": raw.get("project_name")
        if isinstance(raw.get("project_name"), str) and raw.get("project_name")
        else path.parent.name,
        "project_dir": str(path.parent.resolve()),
        "run_mode": raw.get("run_mode") if raw.get("run_mode") in MODES else "guided",
        "run_status": "active",
        "start_at": start_at,
        "stop_after": stop_after,
        "targets": targets,
        "current_stage": None,
        "preferences": {**preferences, "canvas_ratio": "16:9", "offline_ocr": True},
        "policies": {
            "source_policy": preferences.get("source_policy")
            if preferences.get("source_policy") in SOURCE_POLICIES
            else None,
            "data_handling": preferences.get("data_handling")
            if preferences.get("data_handling") in DATA_HANDLING_POLICIES
            else "standard",
            "ocr_policy": "offline",
            "ask_for_ocr_token": False,
        },
        "assumptions": list(raw.get("assumptions", [])) if isinstance(raw.get("assumptions"), list) else [],
        "decisions": list(raw.get("decisions", [])) if isinstance(raw.get("decisions"), list) else [],
        "inputs": {},
        "artifacts": {},
        "deliverables": [],
        "deliverable_artifacts": {},
        "stages": stages,
        "history": [],
        "created_at": raw.get("created_at")
        if isinstance(raw.get("created_at"), str) and raw.get("created_at")
        else now(),
        "updated_at": raw.get("updated_at")
        if isinstance(raw.get("updated_at"), str) and raw.get("updated_at")
        else now(),
    }
    default_authority = raw.get("current_stage") if raw.get("current_stage") in targets else targets[0]
    for role, record in (raw.get("inputs", {}) if isinstance(raw.get("inputs"), dict) else {}).items():
        normalized = normalize_legacy_record(record)
        normalized["authority_stage"] = default_authority
        data["inputs"][role] = normalized

    refresh_dependencies(data)
    for stage in targets:
        item = data["stages"][stage]
        if item["status"] == "completed":
            if stage == "style-options" and not selected_style(data):
                item["status"] = "waiting_user"
            elif stage_completion_errors(data, stage):
                item["status"] = "stale"
                item["qa"] = "not_run"
    for stage in targets:
        item = data["stages"][stage]
        if item["status"] == "completed" and any(
            data["stages"][dependency]["status"] != "completed" for dependency in item["depends_on"]
        ):
            item["status"] = "stale"
            item["qa"] = "not_run"

    legacy_artifacts = raw.get("artifacts", {}) if isinstance(raw.get("artifacts"), dict) else {}
    for key, record in legacy_artifacts.items():
        if key not in DELIVERABLE_STAGE:
            continue
        normalized = normalize_legacy_record(record)
        source_stage = DELIVERABLE_STAGE[key]
        if source_stage not in targets:
            continue
        normalized["stage"] = source_stage
        if data["stages"][source_stage]["status"] != "completed" and normalized.get("status") == "current":
            normalized["status"] = "stale"
        data["deliverables"].append(key)
        data["deliverable_artifacts"][key] = normalized
    data["deliverables"] = canonical_deliverables(data["deliverables"])
    rebuild_artifact_index(data)
    derive_run_state(data, raw.get("run_status") if raw.get("run_status") in {"paused", "blocked", "cancelled"} else None)
    append_history(
        data,
        "migrated",
        details={"from_schema": 1, "backup": str(backup), "added_stage": "deck-revision"},
    )
    return data


def load_locked(path: Path) -> dict[str, Any]:
    if not path.exists():
        raise SystemExit(f"State file not found: {path}")
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise SystemExit(f"Cannot read state file {path}: {exc}") from exc
    version = raw.get("schema_version") if isinstance(raw, dict) else None
    if version == 1:
        backup = backup_manifest(path, "v1")
        data = migrate_v1(raw, path, backup)
        atomic_save(path, data)
        return data
    if version != SCHEMA_VERSION:
        raise SystemExit(f"Unsupported schema_version in {path}: {version!r}")
    validate_manifest(raw, path)
    return raw


def new_stage(status: str) -> dict[str, Any]:
    return {
        "status": status,
        "qa": "not_run",
        "artifacts": {},
        "pages": {},
        "warnings": [],
        "notes": [],
        "depends_on": [],
        "selection": None,
        "revision": 0,
    }


def cmd_init(args: argparse.Namespace) -> None:
    project = Path(args.project_dir).expanduser().resolve()
    with state_lock(project, create=True) as path:
        if path.exists() and not args.force:
            raise SystemExit(f"State file already exists: {path}; use show, retarget, or --force")
        backup = backup_manifest(path, "force") if path.exists() else None
        start_idx = STAGES.index(args.start_at)
        stop_idx = STAGES.index(args.stop_after)
        if stop_idx < start_idx:
            raise SystemExit("--stop-after cannot precede --start-at")
        raw_targets = args.target or default_route_targets(args.start_at, args.stop_after)
        targets = canonical_targets(raw_targets, args.start_at, args.stop_after)
        deliverables = canonical_deliverables(args.deliverable or [])
        validate_deliverable_route(deliverables, targets)
        created = now()
        data: dict[str, Any] = {
            "schema_version": SCHEMA_VERSION,
            "qa_contract_version": 1,
            "state_revision": 0,
            "project_name": args.name or project.name,
            "project_dir": str(project),
            "run_mode": args.mode,
            "run_status": "active",
            "start_at": args.start_at,
            "stop_after": args.stop_after,
            "targets": targets,
            "current_stage": targets[0],
            "preferences": {"canvas_ratio": "16:9", "offline_ocr": True},
            "policies": {
                "source_policy": args.source_policy,
                "data_handling": args.data_handling,
                "ocr_policy": "offline",
                "ask_for_ocr_token": False,
            },
            "assumptions": [],
            "decisions": [],
            "inputs": {},
            "artifacts": {},
            "deliverables": deliverables,
            "deliverable_artifacts": {},
            "stages": {
                stage: new_stage("pending" if stage in targets else "skipped") for stage in STAGES
            },
            "history": [],
            "created_at": created,
            "updated_at": created,
        }
        refresh_dependencies(data)
        append_history(
            data,
            "initialized",
            details={
                "targets": targets,
                "deliverables": deliverables,
                "mode": args.mode,
                "source_policy": args.source_policy,
                "data_handling": args.data_handling,
                "forced_backup": str(backup) if backup else None,
            },
        )
        atomic_save(path, data)
        print(path)


def cmd_show(args: argparse.Namespace) -> None:
    with state_lock(args.project_dir) as path:
        data = load_locked(path)
        if args.json:
            print(json.dumps(data, ensure_ascii=False, indent=2))
            return
        print(f"State: {path}")
        print(f"Project: {data['project_name']} | Schema: v{data['schema_version']} | Revision: {data['state_revision']}")
        print(
            f"Mode: {data['run_mode']} | Run: {data['run_status']} | "
            f"Range: {data['start_at']} -> {data['stop_after']} | Current: {data['current_stage']}"
        )
        print(
            f"Targets: {', '.join(data['targets'])} | "
            f"Deliverables: {', '.join(data['deliverables']) or '(none)'}"
        )
        policies = data.get("policies", {})
        print(
            f"Policies: source={policies.get('source_policy') or '(unset)'} | "
            f"data={policies.get('data_handling', 'standard')} | OCR=offline/no-token"
        )
        for stage in STAGES:
            item = data["stages"][stage]
            print(
                f"- {stage:15} {item['status']:12} qa={item['qa']:20} "
                f"rev={item['revision']} pages={len(item['pages'])}"
            )


def cmd_add_input(args: argparse.Namespace) -> None:
    with state_lock(args.project_dir) as path:
        data = load_locked(path)
        raw_inputs = list(args.inputs or []) + list(args.input or [])
        if not raw_inputs:
            raise SystemExit("add-input requires at least one ROLE=PATH input")
        parsed_inputs = parse_pairs(raw_inputs, "input")
        artifact_role = getattr(args, "artifact_role", None)
        if artifact_role and len(parsed_inputs) != 1:
            raise SystemExit("--artifact-role applies to exactly one imported input per command")
        preferred = data["run_status"] if data["run_status"] in {"paused", "blocked"} else None
        changed_roles: list[str] = []
        for role, raw_path in parsed_inputs.items():
            old = data["inputs"].get(role)
            authority_stage = args.stage or (old or {}).get("authority_stage") or data["current_stage"] or data["targets"][0]
            if authority_stage not in data["targets"]:
                raise SystemExit(
                    f"Input authority stage {authority_stage!r} is not a target; use retarget first"
                )
            ensure_not_self_referential(data, str(raw_path), f"input {role!r}")
            record = file_record(str(raw_path))
            record["authority_stage"] = authority_stage
            imported_role = artifact_role or (old or {}).get("artifact_role")
            if imported_role:
                if authority_stage != "qa-package":
                    raise SystemExit("Imported QA artifacts require --stage qa-package")
                record["artifact_role"] = imported_role
            data["inputs"][role] = record
            metadata_changed = old is not None and any(old.get(key) != record.get(key) for key in ("authority_stage", "artifact_role"))
            if old is None or record_fingerprint(old) != record_fingerprint(record) or metadata_changed:
                changed_roles.append(role)
                if old is not None or data["stages"][authority_stage]["status"] == "completed":
                    invalidate_from(
                        data,
                        authority_stage,
                        include_stage=True,
                        reason=f"input {role!r} changed",
                    )
                    if old and old.get("authority_stage") != authority_stage:
                        invalidate_from(data, old["authority_stage"], include_stage=True, reason=f"input {role!r} reassigned")
        derive_run_state(data, preferred)
        rebuild_artifact_index(data)
        append_history(
            data,
            "inputs-recorded",
            details={"roles": list(parsed_inputs), "changed": changed_roles},
        )
        atomic_save(path, data)
        print(path)


def cmd_record(args: argparse.Namespace) -> None:
    with state_lock(args.project_dir) as path:
        data = load_locked(path)
        preferences = parse_pairs(args.preference, "--preference")
        reserved_policy_keys = {"source_policy", "data_handling", "ocr_policy", "ask_for_ocr_token"}
        conflicting_policy_keys = sorted(reserved_policy_keys.intersection(preferences))
        if conflicting_policy_keys:
            raise SystemExit(
                "Policy keys cannot be set through --preference; use the dedicated policy flags: "
                + ", ".join(conflicting_policy_keys)
            )
        policies = data.setdefault(
            "policies",
            {
                "source_policy": None,
                "data_handling": "standard",
                "ocr_policy": "offline",
                "ask_for_ocr_token": False,
            },
        )
        old_source_policy = policies.get("source_policy")
        old_data_handling = policies.get("data_handling", "standard")
        if args.source_policy is not None:
            policies["source_policy"] = args.source_policy
            preferences["source_policy"] = args.source_policy
        if args.data_handling is not None:
            policies["data_handling"] = args.data_handling
            preferences["data_handling"] = args.data_handling
        policies["ocr_policy"] = "offline"
        policies["ask_for_ocr_token"] = False
        old_style = selected_style(data)
        old_duration = data["preferences"].get("requested_duration_seconds")
        if "requested_duration_seconds" in preferences:
            try:
                timing_number(preferences["requested_duration_seconds"], "requested_duration_seconds", 1, 86400)
            except ValueError as exc:
                raise SystemExit(str(exc)) from exc
        data["preferences"].update(preferences)
        if "requested_duration_seconds" in preferences and old_duration != preferences["requested_duration_seconds"]:
            affected = next((stage for stage in ("speaker-notes", "qa-package") if stage in data["targets"]), None)
            if affected and (data["stages"][affected]["status"] != "pending" or data["stages"][affected]["artifacts"] or data["stages"][affected]["pages"]):
                invalidate_from(data, affected, include_stage=True, reason="user-requested speaking duration changed")
        if "selected_style" in preferences:
            data["stages"]["style-options"]["selection"] = preferences["selected_style"]
            if old_style not in {None, ""} and old_style != preferences["selected_style"]:
                invalidate_from(
                    data,
                    "style-options",
                    include_stage=False,
                    reason="selected visual style changed",
                )
        if args.source_policy is not None and old_source_policy not in {None, args.source_policy}:
            authority = "report" if "report" in data["targets"] else data["targets"][0]
            invalidate_from(
                data,
                authority,
                include_stage=True,
                reason="source policy changed",
            )
        if args.data_handling is not None and old_data_handling != args.data_handling:
            privacy_sensitive = data["targets"][0]
            invalidate_from(
                data,
                privacy_sensitive,
                include_stage=True,
                reason="data-handling policy changed",
            )
        data["assumptions"].extend(x for x in (args.assumption or []) if x not in data["assumptions"])
        data["decisions"].extend(x for x in (args.decision or []) if x not in data["decisions"])
        derive_run_state(data, args.status)
        append_history(
            data,
            "run-recorded",
            details={
                "requested_status": args.status,
                "preferences": list(preferences),
                "source_policy": policies.get("source_policy"),
                "data_handling": policies.get("data_handling"),
                "assumptions_added": args.assumption or [],
                "decisions_added": args.decision or [],
            },
        )
        rebuild_artifact_index(data)
        atomic_save(path, data)
        print(path)


def _store_artifact(data: dict[str, Any], stage: str, key: str, raw_path: Any) -> bool:
    ensure_not_self_referential(data, str(raw_path), f"artifact {key!r}")
    record = file_record(str(raw_path))
    old = data["stages"][stage]["artifacts"].get(key)
    data["stages"][stage]["artifacts"][key] = record
    return old is None or record_fingerprint(old) != record_fingerprint(record)


def _store_deliverable(data: dict[str, Any], stage: str, key: str, raw_path: Any) -> bool:
    if key not in data["deliverables"]:
        raise SystemExit(f"Deliverable {key!r} is not declared; use retarget --add-deliverable first")
    expected_stage = DELIVERABLE_STAGE[key]
    if stage != expected_stage:
        raise SystemExit(f"Deliverable {key!r} belongs to stage {expected_stage!r}, not {stage!r}")
    ensure_not_self_referential(data, str(raw_path), f"deliverable {key!r}")
    record = file_record(str(raw_path))
    record["stage"] = stage
    old = data["deliverable_artifacts"].get(key)
    data["deliverable_artifacts"][key] = record
    return old is None or record_fingerprint(old) != record_fingerprint(record) or old.get("stage") != stage


def cmd_set_stage(args: argparse.Namespace) -> None:
    with state_lock(args.project_dir) as path:
        data = load_locked(path)
        if args.stage not in data["targets"]:
            has_mutation = bool(
                args.qa
                or args.note
                or args.warning
                or args.selected_style is not None
                or args.artifact
                or args.deliverable
                or args.status not in {None, "skipped"}
            )
            if has_mutation:
                raise SystemExit(f"Stage {args.stage!r} is not a target; use retarget before recording work")
        item = data["stages"][args.stage]
        old_status = item["status"]
        authority_changed = False
        if args.qa:
            item["qa"] = args.qa
        if args.note:
            item["notes"].append(args.note)
        for warning in args.warning or []:
            if warning not in item["warnings"]:
                item["warnings"].append(warning)
        if args.selected_style is not None:
            old_style = selected_style(data)
            item["selection"] = args.selected_style
            data["preferences"]["selected_style"] = args.selected_style
            authority_changed = old_style not in {None, ""} and old_style != args.selected_style
        for key, raw_path in parse_pairs(args.artifact, "--artifact").items():
            authority_changed = _store_artifact(data, args.stage, key, raw_path) or authority_changed
        for key, raw_path in parse_pairs(args.deliverable, "--deliverable").items():
            authority_changed = _store_deliverable(data, args.stage, key, raw_path) or authority_changed

        requested_status = args.status
        guided_wait = (
            requested_status == "completed"
            and args.stage == "style-options"
            and data["run_mode"] in {"guided", "stepwise"}
            and not selected_style(data)
            and bool({"template", "image-deck", "editable-deck"}.intersection(data.get("targets", [])))
        )
        if guided_wait:
            item["status"] = "waiting_user"
        elif requested_status:
            if requested_status == "skipped" and args.stage in data["targets"]:
                raise SystemExit("A target stage cannot be skipped; use retarget to remove it")
            item["status"] = requested_status

        explicitly_reaccepted = requested_status == "completed" and args.qa in PASSING_QA
        if old_status == "completed" and authority_changed and not explicitly_reaccepted:
            item["status"] = "stale"
            item["qa"] = "not_run"

        if item["status"] == "completed":
            drift = detect_drift(data)
            if drift:
                raise SystemExit(
                    f"Cannot complete stage {args.stage!r}; current artifacts drifted after the last verify:\n"
                    + "\n".join(f"- {entry}" for entry in drift)
                    + "\nRun project_state.py verify, repair/re-register authorities, then re-run final QA."
                )
            errors = stage_completion_errors(data, args.stage)
            if errors:
                raise SystemExit(
                    f"Cannot complete stage {args.stage!r}:\n" + "\n".join(f"- {error}" for error in errors)
                )
        if args.qa == "failed" and item["status"] == "completed":
            item["status"] = "failed"

        if old_status == "completed" and (item["status"] != "completed" or authority_changed):
            if item["status"] == "completed":
                errors = stage_completion_errors(data, args.stage)
                if errors:
                    raise SystemExit(
                        f"Cannot re-complete stage {args.stage!r}:\n"
                        + "\n".join(f"- {error}" for error in errors)
                    )
            else:
                item["qa"] = "not_run" if item["status"] in {"pending", "running", "stale"} else item["qa"]
            invalidate_from(
                data,
                args.stage,
                include_stage=False,
                reason=f"upstream stage {args.stage!r} changed",
            )
        elif authority_changed:
            invalidate_from(
                data,
                args.stage,
                include_stage=False,
                reason=f"authority for stage {args.stage!r} changed",
            )

        item["revision"] = int(item.get("revision", 0)) + 1
        if guided_wait:
            preferred = "paused"
        elif requested_status == "completed" and data["run_mode"] == "stepwise":
            preferred = "paused"
        elif requested_status == "waiting_user":
            preferred = "paused"
        elif requested_status == "failed":
            preferred = "blocked"
        elif requested_status == "cancelled":
            preferred = "cancelled"
        elif requested_status in {"running", "pending", "stale", "completed"}:
            preferred = "active"
        else:
            preferred = data["run_status"] if data["run_status"] in {"paused", "blocked"} else None
        derive_run_state(data, preferred)
        rebuild_artifact_index(data)
        append_history(
            data,
            "stage-updated",
            stage=args.stage,
            details={
                "requested_status": requested_status,
                "effective_status": item["status"],
                "qa": item["qa"],
                "guided_wait": guided_wait,
                "authority_changed": authority_changed,
            },
        )
        atomic_save(path, data)
        print(path)


def cmd_invalidate(args: argparse.Namespace) -> None:
    with state_lock(args.project_dir) as path:
        data = load_locked(path)
        affected = invalidate_from(
            data,
            args.from_stage,
            include_stage=True,
            reason=args.reason,
        )
        derive_run_state(data, "active")
        append_history(
            data,
            "invalidated",
            stage=args.from_stage,
            details={"affected": affected, "reason": args.reason},
        )
        atomic_save(path, data)
        print(path)


def cmd_retarget(args: argparse.Namespace) -> None:
    with state_lock(args.project_dir) as path:
        data = load_locked(path)
        if not any(
            [
                args.start_at,
                args.stop_after,
                args.target,
                args.mode,
                args.deliverable,
                args.add_deliverable,
                args.remove_deliverable,
                args.clear_deliverables,
            ]
        ):
            raise SystemExit(
                "retarget requires a route, mode, or deliverable change"
            )
        old_targets = list(data["targets"])
        old_deliverables = list(data["deliverables"])
        start_at = args.start_at or data["start_at"]
        stop_after = args.stop_after or data["stop_after"]
        start_idx = STAGES.index(start_at)
        stop_idx = STAGES.index(stop_after)
        if stop_idx < start_idx:
            raise SystemExit("--stop-after cannot precede --start-at")
        if args.target:
            targets = canonical_targets(args.target, start_at, stop_after)
        elif args.start_at or args.stop_after:
            targets = default_route_targets(start_at, stop_after)
        else:
            targets = old_targets
        if args.clear_deliverables and args.deliverable:
            raise SystemExit("Use either --clear-deliverables or --deliverable, not both")
        if args.clear_deliverables:
            deliverables: list[str] = []
        elif args.deliverable:
            deliverables = canonical_deliverables(args.deliverable)
        else:
            deliverables = list(old_deliverables)
        deliverable_set = set(deliverables)
        deliverable_set.update(args.add_deliverable or [])
        deliverable_set.difference_update(args.remove_deliverable or [])
        deliverables = canonical_deliverables(list(deliverable_set))
        validate_deliverable_route(deliverables, targets)
        data["start_at"] = start_at
        data["stop_after"] = stop_after
        data["targets"] = targets
        data["deliverables"] = deliverables
        if args.mode:
            data["run_mode"] = args.mode
        remapped_inputs: dict[str, str] = {}
        for role, record in data["inputs"].items():
            authority_stage = record.get("authority_stage")
            if authority_stage not in targets:
                old_index = STAGES.index(authority_stage) if authority_stage in STAGES else -1
                replacement = next(
                    (candidate for candidate in targets if STAGES.index(candidate) >= old_index),
                    targets[0],
                )
                record["authority_stage"] = replacement
                if replacement != "qa-package" and "artifact_role" in record:
                    record["previous_artifact_role"] = record.pop("artifact_role")
                remapped_inputs[role] = replacement
        refresh_dependencies(data)
        for stage in STAGES:
            item = data["stages"][stage]
            if stage not in targets:
                item["status"] = "skipped"
                for record in item.get("artifacts", {}).values():
                    if isinstance(record, dict) and record.get("status") == "current":
                        record["status"] = "unrequested"
                for page in item.get("pages", {}).values():
                    if not isinstance(page, dict):
                        continue
                    for record in page.get("artifacts", {}).values():
                        if isinstance(record, dict) and record.get("status") == "current":
                            record["status"] = "unrequested"
                for deliverable in data["deliverable_artifacts"].values():
                    if deliverable.get("stage") == stage and deliverable.get("status") == "current":
                        deliverable["status"] = "unrequested"
            elif stage not in old_targets:
                if item["status"] in {"skipped", "cancelled"}:
                    item["status"] = "pending"
                    item["qa"] = "not_run"
                for deliverable in data["deliverable_artifacts"].values():
                    if deliverable.get("stage") == stage and deliverable.get("status") == "unrequested":
                        deliverable["status"] = "stale"
                for record in item.get("artifacts", {}).values():
                    if isinstance(record, dict) and record.get("status") == "unrequested":
                        record["status"] = "stale"
                for page in item.get("pages", {}).values():
                    if not isinstance(page, dict):
                        continue
                    for record in page.get("artifacts", {}).values():
                        if isinstance(record, dict) and record.get("status") == "unrequested":
                            record["status"] = "stale"
        for role, artifact in data["deliverable_artifacts"].items():
            if role not in deliverables and artifact.get("status") == "current":
                artifact["status"] = "unrequested"
            elif role in deliverables and artifact.get("status") == "unrequested":
                artifact["status"] = "stale"
        for stage in targets:
            item = data["stages"][stage]
            if item["status"] == "completed" and stage_completion_errors(data, stage):
                mark_stage_stale(data, stage, "route dependencies changed")
        drift = detect_drift(data)
        if drift:
            derive_run_state(data, "active")
        derive_run_state(data, "active")
        rebuild_artifact_index(data)
        append_history(
            data,
            "retargeted",
            details={
                "old_targets": old_targets,
                "targets": targets,
                "old_deliverables": old_deliverables,
                "deliverables": deliverables,
                "mode": data["run_mode"],
                "remapped_inputs": remapped_inputs,
            },
        )
        atomic_save(path, data)
        print(path)


def cmd_pause(args: argparse.Namespace) -> None:
    with state_lock(args.project_dir) as path:
        data = load_locked(path)
        if first_unresolved_target(data) is None:
            raise SystemExit("Completed run cannot be paused")
        derive_run_state(data, "paused")
        append_history(data, "paused", details={"reason": args.reason})
        atomic_save(path, data)
        print(path)


def _mark_record_drift(record: dict[str, Any], observed: dict[str, Any]) -> None:
    record["status"] = "missing" if not observed.get("exists") else "drifted"
    record["observed"] = observed
    record["verified_at"] = now()


def detect_drift(data: dict[str, Any]) -> list[dict[str, Any]]:
    drift: list[dict[str, Any]] = []
    for role, record in data["inputs"].items():
        observed = file_record(record["path"])
        if record_fingerprint(record) != record_fingerprint(observed):
            authority_stage = record.get("authority_stage") or data["targets"][0]
            invalidate_from(data, authority_stage, include_stage=True, reason=f"input {role!r} drifted")
            _mark_record_drift(record, observed)
            drift.append({"kind": "input", "key": role, "stage": authority_stage, "path": record["path"]})
        else:
            record["verified_at"] = now()

    for stage in STAGES:
        item = data["stages"][stage]
        for key, record in item["artifacts"].items():
            if record.get("status") == "unrequested":
                continue
            observed = file_record(record["path"])
            if record_fingerprint(record) != record_fingerprint(observed):
                invalidate_from(data, stage, include_stage=True, reason=f"artifact {stage}:{key} drifted")
                _mark_record_drift(record, observed)
                drift.append({"kind": "artifact", "key": key, "stage": stage, "path": record["path"]})
            else:
                record["verified_at"] = now()
        for page_id, page in item["pages"].items():
            for key, record in page["artifacts"].items():
                if record.get("status") == "unrequested":
                    continue
                observed = file_record(record["path"])
                if record_fingerprint(record) != record_fingerprint(observed):
                    invalidate_from(data, stage, include_stage=True, reason=f"page artifact {stage}:{page_id}:{key} drifted")
                    page["status"] = "stale"
                    page["qa"] = "not_run"
                    _mark_record_drift(record, observed)
                    drift.append(
                        {"kind": "page-artifact", "key": key, "stage": stage, "page": page_id, "path": record["path"]}
                    )
                else:
                    record["verified_at"] = now()

    for key, record in data["deliverable_artifacts"].items():
        if record.get("status") == "unrequested":
            continue
        observed = file_record(record["path"])
        if record_fingerprint(record) != record_fingerprint(observed):
            stage = record["stage"]
            invalidate_from(data, stage, include_stage=True, reason=f"deliverable {key!r} drifted")
            _mark_record_drift(record, observed)
            drift.append({"kind": "deliverable", "key": key, "stage": stage, "path": record["path"]})
        else:
            record["verified_at"] = now()
    rebuild_artifact_index(data)
    return drift


def cmd_verify(args: argparse.Namespace) -> None:
    with state_lock(args.project_dir) as path:
        data = load_locked(path)
        previous = data["run_status"]
        drift = detect_drift(data)
        preferred = previous if previous in {"paused", "blocked", "cancelled"} else "active"
        derive_run_state(data, preferred)
        append_history(data, "verified", details={"drift_count": len(drift), "drift": drift})
        atomic_save(path, data)
        result = {
            "passed": not drift,
            "drift_count": len(drift),
            "drift": drift,
            "run_status": data["run_status"],
            "current_stage": data["current_stage"],
            "state_revision": data["state_revision"],
        }
        if args.json:
            print(json.dumps(result, ensure_ascii=False, indent=2))
        else:
            print(f"Verification: {'passed' if not drift else 'drift detected'} | count={len(drift)}")
            for item in drift:
                print(f"- {item['kind']} {item['key']} ({item['path']})")
        if drift and not args.no_fail:
            raise SystemExit(2)


def cmd_resume(args: argparse.Namespace) -> None:
    with state_lock(args.project_dir) as path:
        data = load_locked(path)
        if data["run_status"] == "cancelled" and not args.force:
            raise SystemExit("Cancelled run requires resume --force")
        drift = detect_drift(data)
        for stage in data["targets"]:
            item = data["stages"][stage]
            if item["status"] == "cancelled" and args.force:
                item["status"] = "pending"
                item["qa"] = "not_run"
            elif item["status"] == "failed" and args.retry_failed:
                item["status"] = "pending"
                item["qa"] = "not_run"
            elif item["status"] == "running" and args.restart_running:
                item["status"] = "pending"
                item["qa"] = "not_run"
            elif item["status"] == "waiting_user":
                if stage == "style-options" and selected_style(data):
                    item["status"] = "completed"
                    errors = stage_completion_errors(data, stage)
                    if errors:
                        item["status"] = "pending"
                        item["qa"] = "not_run"
                elif args.force:
                    item["status"] = "pending"
                    item["qa"] = "not_run"
        derive_run_state(data, "active")
        rebuild_artifact_index(data)
        append_history(
            data,
            "resumed",
            details={
                "retry_failed": args.retry_failed,
                "restart_running": args.restart_running,
                "force": args.force,
                "drift_count": len(drift),
            },
        )
        atomic_save(path, data)
        result = {
            "run_status": data["run_status"],
            "current_stage": data["current_stage"],
            "drift_count": len(drift),
            "state_revision": data["state_revision"],
        }
        if args.json:
            print(json.dumps(result, ensure_ascii=False, indent=2))
        else:
            print(
                f"Run: {data['run_status']} | Current: {data['current_stage']} | "
                f"Drift detected: {len(drift)}"
            )


def cmd_adopt(args: argparse.Namespace) -> None:
    with state_lock(args.project_dir) as path:
        data = load_locked(path)
        if args.stage not in data["targets"]:
            raise SystemExit(f"Stage {args.stage!r} is not a target; use retarget first")
        inputs = parse_pairs(args.input, "--input")
        artifacts = parse_pairs(args.artifact, "--artifact")
        deliverable_artifacts = parse_pairs(args.deliverable, "--deliverable")
        if not inputs and not artifacts and not deliverable_artifacts and args.selected_style is None:
            raise SystemExit("adopt requires --input, --artifact, --deliverable, or --selected-style")
        item = data["stages"][args.stage]
        for role, raw_path in inputs.items():
            ensure_not_self_referential(data, str(raw_path), f"input {role!r}")
            record = ensure_existing_record(str(raw_path), f"input {role!r}")
            record.update({"authority_stage": args.stage, "adopted": True})
            previous_role = data["inputs"].get(role, {}).get("artifact_role")
            if previous_role:
                record["artifact_role" if args.stage == "qa-package" else "previous_artifact_role"] = previous_role
            data["inputs"][role] = record
        for key, raw_path in artifacts.items():
            ensure_not_self_referential(data, str(raw_path), f"artifact {key!r}")
            record = ensure_existing_record(str(raw_path), f"artifact {key!r}")
            record["adopted"] = True
            item["artifacts"][key] = record
        for key, raw_path in deliverable_artifacts.items():
            if key not in data["deliverables"]:
                raise SystemExit(
                    f"Deliverable {key!r} is not declared; use retarget --add-deliverable first"
                )
            if DELIVERABLE_STAGE[key] != args.stage:
                raise SystemExit(
                    f"Deliverable {key!r} belongs to stage {DELIVERABLE_STAGE[key]!r}, not {args.stage!r}"
                )
            ensure_not_self_referential(data, str(raw_path), f"deliverable {key!r}")
            record = ensure_existing_record(str(raw_path), f"deliverable {key!r}")
            record.update({"stage": args.stage, "adopted": True})
            data["deliverable_artifacts"][key] = record
        if args.selected_style is not None:
            item["selection"] = args.selected_style
            data["preferences"]["selected_style"] = args.selected_style
        if args.note:
            item["notes"].append(args.note)
        invalidate_from(
            data,
            args.stage,
            include_stage=False,
            reason=f"authority for stage {args.stage!r} was adopted",
        )
        if args.complete:
            if args.qa not in PASSING_QA:
                raise SystemExit("adopt --complete requires --qa passed or passed_with_warnings")
            item["qa"] = args.qa
            item["status"] = "completed"
            drift = detect_drift(data)
            if drift:
                raise SystemExit(
                    f"Cannot complete adopted stage {args.stage!r}; current artifacts drifted:\n"
                    + "\n".join(f"- {entry}" for entry in drift)
                )
            errors = stage_completion_errors(data, args.stage)
            if errors:
                raise SystemExit(
                    f"Cannot complete adopted stage {args.stage!r}:\n"
                    + "\n".join(f"- {error}" for error in errors)
                )
        elif item["status"] == "completed":
            item["status"] = "stale"
            item["qa"] = args.qa or "not_run"
        elif args.qa:
            item["qa"] = args.qa
        item["revision"] = int(item.get("revision", 0)) + 1
        preferred = "paused" if args.complete and data["run_mode"] == "stepwise" else "active"
        derive_run_state(data, preferred)
        rebuild_artifact_index(data)
        append_history(
            data,
            "adopted",
            stage=args.stage,
            details={
                "inputs": list(inputs),
                "artifacts": list(artifacts),
                "deliverables": list(deliverable_artifacts),
                "complete": args.complete,
            },
        )
        atomic_save(path, data)
        print(path)


def cmd_set_page(args: argparse.Namespace) -> None:
    with state_lock(args.project_dir) as path:
        data = load_locked(path)
        if args.stage not in data["targets"]:
            raise SystemExit(f"Stage {args.stage!r} is not a target; use retarget first")
        stage_item = data["stages"][args.stage]
        page_id = str(args.page)
        page = stage_item["pages"].setdefault(
            page_id,
            {
                "status": "pending",
                "qa": "not_run",
                "artifacts": {},
                "warnings": [],
                "notes": [],
                "revision": 0,
            },
        )
        stage_was_completed = stage_item["status"] == "completed"
        if args.qa:
            page["qa"] = args.qa
        if args.note:
            page["notes"].append(args.note)
        for warning in args.warning or []:
            if warning not in page["warnings"]:
                page["warnings"].append(warning)
        artifact_changed = False
        for key, raw_path in parse_pairs(args.artifact, "--artifact").items():
            ensure_not_self_referential(data, str(raw_path), f"page artifact {key!r}")
            record = file_record(str(raw_path))
            old = page["artifacts"].get(key)
            page["artifacts"][key] = record
            artifact_changed = artifact_changed or old is None or record_fingerprint(old) != record_fingerprint(record)
        if args.status:
            page["status"] = args.status
        if page["status"] == "completed":
            drift = detect_drift(data)
            if drift:
                raise SystemExit(
                    f"Cannot complete page {page_id!r}; current project artifacts drifted:\n"
                    + "\n".join(f"- {entry}" for entry in drift)
                )
            errors = page_completion_errors(page_id, page)
            if errors:
                raise SystemExit(
                    f"Cannot complete page {page_id!r}:\n" + "\n".join(f"- {error}" for error in errors)
                )
        if stage_was_completed and (artifact_changed or page["status"] != "completed"):
            mark_stage_stale(data, args.stage, f"page {page_id!r} changed")
            invalidate_from(
                data,
                args.stage,
                include_stage=False,
                reason=f"page {page_id!r} changed",
            )
        elif page["status"] == "failed":
            stage_item["status"] = "failed"
            stage_item["qa"] = "failed"
            invalidate_from(data, args.stage, include_stage=False, reason=f"page {page_id!r} failed")
        elif page["status"] in {"running", "waiting_user"} and stage_item["status"] in {"pending", "stale"}:
            stage_item["status"] = page["status"]
        page["revision"] = int(page.get("revision", 0)) + 1
        stage_item["revision"] = int(stage_item.get("revision", 0)) + 1
        if page["status"] == "waiting_user":
            preferred = "paused"
        elif page["status"] == "failed":
            preferred = "blocked"
        else:
            preferred = "active"
        derive_run_state(data, preferred)
        rebuild_artifact_index(data)
        append_history(
            data,
            "page-updated",
            stage=args.stage,
            details={"page": page_id, "status": page["status"], "qa": page["qa"]},
        )
        atomic_save(path, data)
        print(path)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest="command", required=True)

    init = sub.add_parser("init", help="create a schema-v2 ppt-gen-state.json")
    init.add_argument("project_dir")
    init.add_argument("--name")
    init.add_argument("--mode", choices=sorted(MODES), default="guided")
    init.add_argument("--source-policy", choices=sorted(SOURCE_POLICIES), default="verify-update")
    init.add_argument(
        "--data-handling", choices=sorted(DATA_HANDLING_POLICIES), default="standard"
    )
    init.add_argument("--start-at", choices=STAGES, default="brief")
    init.add_argument("--stop-after", choices=STAGES, default="qa-package")
    init.add_argument("--target", action="append", choices=STAGES)
    init.add_argument("--deliverable", action="append", choices=DELIVERABLE_ROLES)
    init.add_argument("--force", action="store_true")
    init.set_defaults(func=cmd_init)

    show = sub.add_parser("show", help="show and structurally validate project state")
    show.add_argument("project_dir")
    show.add_argument("--json", action="store_true")
    show.set_defaults(func=cmd_show)

    add_input = sub.add_parser("add-input", help="record role=path inputs and recursive hashes")
    add_input.add_argument("project_dir")
    add_input.add_argument("inputs", nargs="*", metavar="ROLE=PATH")
    add_input.add_argument("--input", action="append", help="backward-compatible ROLE=PATH form")
    add_input.add_argument("--stage", choices=STAGES, help="stage whose authority this input satisfies")
    add_input.add_argument("--artifact-role", choices=sorted(CLI_DELIVERABLE_ROLES), help="Explicit imported artifact to audit at qa-package; not a new production deliverable")
    add_input.set_defaults(func=cmd_add_input)

    record = sub.add_parser("record", help="record run-level assumptions, decisions, and preferences")
    record.add_argument("project_dir")
    record.add_argument("--status", choices=sorted(RUN_STATES))
    record.add_argument("--source-policy", choices=sorted(SOURCE_POLICIES))
    record.add_argument("--data-handling", choices=sorted(DATA_HANDLING_POLICIES))
    record.add_argument("--preference", action="append")
    record.add_argument("--assumption", action="append")
    record.add_argument("--decision", action="append")
    record.set_defaults(func=cmd_record)

    set_stage = sub.add_parser("set-stage", help="update a phase and its artifacts")
    set_stage.add_argument("project_dir")
    set_stage.add_argument("stage", choices=STAGES)
    set_stage.add_argument("--status", choices=sorted(STAGE_STATES))
    set_stage.add_argument("--qa", choices=sorted(QA_STATES))
    set_stage.add_argument("--artifact", action="append")
    set_stage.add_argument("--deliverable", action="append")
    set_stage.add_argument("--warning", action="append")
    set_stage.add_argument("--note")
    set_stage.add_argument("--selected-style")
    set_stage.set_defaults(func=cmd_set_stage)

    invalidate = sub.add_parser("invalidate", help="mark a phase and its dependency consumers stale")
    invalidate.add_argument("project_dir")
    invalidate.add_argument("--from-stage", choices=STAGES, required=True)
    invalidate.add_argument("--reason")
    invalidate.set_defaults(func=cmd_invalidate)

    verify = sub.add_parser("verify", help="rehash inputs/artifacts and invalidate drifted consumers")
    verify.add_argument("project_dir")
    verify.add_argument("--json", action="store_true")
    verify.add_argument("--no-fail", action="store_true", help="exit zero even when drift is detected")
    verify.set_defaults(func=cmd_verify)

    retarget = sub.add_parser("retarget", help="change requested range, targets, or run mode without resetting history")
    retarget.add_argument("project_dir")
    retarget.add_argument("--start-at", choices=STAGES)
    retarget.add_argument("--stop-after", choices=STAGES)
    retarget.add_argument("--target", action="append", choices=STAGES)
    retarget.add_argument("--mode", choices=sorted(MODES))
    retarget.add_argument(
        "--deliverable",
        action="append",
        choices=DELIVERABLE_ROLES,
        help="replace the requested deliverable set",
    )
    retarget.add_argument("--add-deliverable", action="append", choices=DELIVERABLE_ROLES)
    retarget.add_argument("--remove-deliverable", action="append", choices=DELIVERABLE_ROLES)
    retarget.add_argument("--clear-deliverables", action="store_true")
    retarget.set_defaults(func=cmd_retarget)

    pause = sub.add_parser("pause", help="pause the run at its derived current target")
    pause.add_argument("project_dir")
    pause.add_argument("--reason")
    pause.set_defaults(func=cmd_pause)

    resume = sub.add_parser("resume", help="verify and resume the first unresolved target")
    resume.add_argument("project_dir")
    resume.add_argument("--retry-failed", action="store_true")
    resume.add_argument("--restart-running", action="store_true")
    resume.add_argument("--force", action="store_true", help="resume cancelled or user-waiting work")
    resume.add_argument("--json", action="store_true")
    resume.set_defaults(func=cmd_resume)

    adopt = sub.add_parser("adopt", help="adopt user-edited authority without overwriting it")
    adopt.add_argument("project_dir")
    adopt.add_argument("stage", choices=STAGES)
    adopt.add_argument("--input", action="append")
    adopt.add_argument("--artifact", action="append")
    adopt.add_argument("--deliverable", action="append")
    adopt.add_argument("--selected-style")
    adopt.add_argument("--qa", choices=sorted(QA_STATES))
    adopt.add_argument("--complete", action="store_true")
    adopt.add_argument("--note")
    adopt.set_defaults(func=cmd_adopt)

    set_page = sub.add_parser("set-page", help="record per-page progress and artifacts")
    set_page.add_argument("project_dir")
    set_page.add_argument("stage", choices=STAGES)
    set_page.add_argument("page")
    set_page.add_argument("--status", choices=sorted(PAGE_STATES))
    set_page.add_argument("--qa", choices=sorted(QA_STATES))
    set_page.add_argument("--artifact", action="append")
    set_page.add_argument("--warning", action="append")
    set_page.add_argument("--note")
    set_page.set_defaults(func=cmd_set_page)
    return parser


def main() -> None:
    args = build_parser().parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
