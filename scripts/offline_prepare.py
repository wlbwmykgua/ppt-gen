#!/usr/bin/env python3
"""Prepare approved slide images with guarded editppt and local-only OCR hints."""
from __future__ import annotations

import argparse
import json
import os
import subprocess
from pathlib import Path

from local_runtime import configure_cache, locate_tool, recognize_image, sha256


def binding(path: Path) -> dict:
    return {"path": str(path.resolve()), "sha256": sha256(path)}


def check_binding(record: dict) -> Path:
    if not isinstance(record, dict) or not isinstance(record.get("path"), str):
        raise ValueError("a path/SHA-256 binding is required")
    path = Path(record["path"]).expanduser()
    if not path.is_absolute() or not path.is_file() or sha256(path) != record.get("sha256"):
        raise ValueError(f"missing or hash-mismatched bound file: {path}")
    return path.resolve()


def prepare_command(handoff: dict, job_dir: Path) -> list[str]:
    if not isinstance(handoff, dict):
        raise ValueError("handoff must be an object")
    if handoff.get("schema_version") != 1 or handoff.get("ocr_policy") != "offline" or handoff.get("ask_for_ocr_token") is not False:
        raise ValueError("handoff must require schema_version=1 and offline/no-token OCR")
    if handoff.get("data_handling") == "local-only":
        # Current editppt cannot record a genuinely disabled image backend.
        # Do not silently configure its remote-capable CLI fallback instead.
        raise ValueError("installed editppt has no disabled image-backend contract; local-only reconstruction requires a compatible backend before prepare")
    if handoff.get("data_handling") != "standard" or handoff.get("image_backend") != "builtin-allowed":
        raise ValueError("unsupported data-handling/image-backend policy")
    for key in ("source_manifest", "slide_copy_ledger", "slide_manifest"):
        check_binding({"path": handoff.get(key), "sha256": handoff.get("hashes", {}).get(key) or handoff.get(key + "_sha256")})
    images = handoff.get("source_images")
    if not isinstance(images, list) or not images:
        raise ValueError("handoff requires ordered source_images")
    paths = [check_binding(record) for record in images]
    executable = locate_tool("editppt")
    if not executable:
        raise ValueError("editppt is not installed or PPT_GEN_EDITPPT is invalid")
    return [executable, "prepare", "--job-dir", str(job_dir.resolve()),
            "--no-text-hints", "--image-backend", "builtin-imagegen", *map(str, paths)]


def prepare(handoff_path: Path, job_dir: Path, report_path: Path, dry_run: bool = False) -> dict:
    handoff_path, job_dir, report_path = (p.expanduser().resolve() for p in (handoff_path, job_dir, report_path))
    handoff = json.loads(handoff_path.read_text(encoding="utf-8"))
    command = prepare_command(handoff, job_dir)
    if dry_run:
        return {"command": command, "ocr_policy": "offline", "network_calls": 0}
    if job_dir.exists() or report_path.exists():
        raise ValueError("run/report already exists; preserve it and select a new versioned destination")
    help_result = subprocess.run([command[0], "prepare", "--help"], capture_output=True, text=True, check=True, timeout=30)
    if "--no-text-hints" not in help_result.stdout:
        raise ValueError("editppt lacks --no-text-hints; refusing the unguarded prepare command")
    environment = dict(os.environ)
    environment.pop("PADDLE_OCR_TOKEN", None)
    # The explicit flag is essential: clearing the environment alone would
    # still allow editppt's separate credential-file fallback.
    subprocess.run(command, env=environment, capture_output=True, text=True, check=True, timeout=120)
    jobs = json.loads((job_dir / "page_jobs.json").read_text(encoding="utf-8"))
    pages = jobs.get("pages", [])
    if len(pages) != len(handoff["source_images"]):
        raise ValueError("prepared page count differs from handoff")
    evidence = []
    for index, page in enumerate(pages, 1):
        page_dir = (job_dir / page["page_dir"]).resolve()
        if not page_dir.is_relative_to(job_dir / "pages"):
            raise ValueError("prepared page escapes its run directory")
        source = page_dir / "source.png"
        subprocess.run([command[0], "page", "hints", str(page_dir)], env=environment,
                       capture_output=True, text=True, check=True, timeout=120)
        recognition = recognize_image(source)
        transcript = page_dir / "offline-recognition.json"
        transcript.write_text(json.dumps({**recognition, "source": binding(source)}, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        evidence.append({"page": index, "source": binding(source),
                         "geometry_hints": binding(page_dir / "text_hints.json"),
                         "recognition": binding(transcript)})
    report = {"schema_version": 1, "tool": "ppt-gen.offline-prepare", "passed": True,
              "handoff": binding(handoff_path), "source_images": handoff["source_images"],
              "prepare_command": command, "ocr_policy": "offline", "external_ocr_used": False,
              "backend": "apple-vision-offline", "pages": evidence}
    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return {"report": str(report_path), "run": str(job_dir), "pages": len(pages), "ocr_policy": "offline"}


def validate_report(path: Path, handoff_path: Path | None) -> list[str]:
    """Bind execution provenance and local recognition files to current inputs."""
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        if data.get("schema_version") != 1 or data.get("tool") != "ppt-gen.offline-prepare" or data.get("passed") is not True:
            raise ValueError("invalid offline-preparation report")
        if data.get("ocr_policy") != "offline" or data.get("external_ocr_used") is not False or data.get("backend") != "apple-vision-offline":
            raise ValueError("offline-preparation backend/policy mismatch")
        bound_handoff = check_binding(data.get("handoff"))
        if handoff_path is None or bound_handoff != handoff_path.resolve():
            raise ValueError("offline-preparation binds a different handoff")
        handoff = json.loads(bound_handoff.read_text(encoding="utf-8"))
        sources = data.get("source_images")
        if not isinstance(sources, list) or not sources or sources != handoff.get("source_images"):
            raise ValueError("offline-preparation source order differs from handoff")
        source_paths = [str(check_binding(source)) for source in sources]
        command = data.get("prepare_command", [])
        if (not isinstance(command, list) or len(command) != 7 + len(sources)
                or command[1:3] != ["prepare", "--job-dir"]
                or command[4:7] != ["--no-text-hints", "--image-backend", "builtin-imagegen"]
                or command[7:] != source_paths):
            raise ValueError("offline-preparation lacks the guarded prepare command")
        pages = data.get("pages", [])
        if len(pages) != len(sources):
            raise ValueError("offline-preparation page count mismatch")
        for index, page in enumerate(pages, 1):
            if page.get("page") != index:
                raise ValueError("offline-preparation pages must be contiguous")
            source = check_binding(page.get("source"))
            hint_path = check_binding(page.get("geometry_hints"))
            recognition_path = check_binding(page.get("recognition"))
            if not source.is_relative_to(Path(command[3]).resolve() / "pages") or hint_path.parent != source.parent or recognition_path.parent != source.parent:
                raise ValueError("offline page evidence escapes its prepared page directory")
            from PIL import Image
            with Image.open(source) as normalized, Image.open(source_paths[index - 1]) as original:
                if normalized.size != original.size or normalized.convert("RGBA").tobytes() != original.convert("RGBA").tobytes():
                    raise ValueError("prepared source pixels differ from the corresponding approved slide")
            hints = json.loads(hint_path.read_text(encoding="utf-8"))
            if hints.get("backend") not in {None, "builtin-ink"}:
                raise ValueError("geometry hints were not produced by the local detector")
            recognition = json.loads(recognition_path.read_text(encoding="utf-8"))
            if recognition.get("backend") != "apple-vision-offline" or not isinstance(recognition.get("lines"), list):
                raise ValueError("offline text recognition is missing")
            if recognition.get("source") != page.get("source"):
                raise ValueError("recognition source mismatch")
        return []
    except (OSError, ValueError, KeyError, TypeError, AttributeError, ImportError) as exc:
        return [f"offline-preparation: {exc}"]


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--handoff", type=Path, required=True)
    parser.add_argument("--job-dir", type=Path, required=True)
    parser.add_argument("--report", type=Path, required=True)
    parser.add_argument("--cache-dir", type=Path)
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()
    configure_cache(args.cache_dir)
    try:
        output = prepare(args.handoff, args.job_dir, args.report, args.dry_run)
    except (ValueError, OSError, subprocess.SubprocessError, RuntimeError, KeyError, TypeError, AttributeError) as exc:
        raise SystemExit(f"offline prepare stopped: {exc}") from exc
    print(json.dumps(output, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
