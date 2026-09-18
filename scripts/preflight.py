#!/usr/bin/env python3
"""Run ppt-gen environment and privacy preflight checks with JSON output.

Exit codes are intentionally tri-state:
  0 = passed
  2 = passed_with_warnings
  1 = failed
"""

from __future__ import annotations

import argparse
import importlib.util
import json
import os
import platform
import re
import shutil
import subprocess
import sys
from datetime import datetime
from pathlib import Path
from typing import Any

from local_runtime import locate_tool, font_roots


STATUSES = ("passed", "passed_with_warnings", "failed")
STAGES = (
    "brief",
    "report",
    "outline",
    "style-options",
    "template",
    "image-deck",
    "editable-deck",
    "speaker-notes",
    "qa-package",
    "deck-revision",
)
FONT_EXTENSIONS = {".ttf", ".otf", ".ttc", ".dfont"}
DEFAULT_APPS = ("Microsoft PowerPoint", "Keynote", "LibreOffice", "WPS Office")
APP_ALIASES = {
    "microsoft powerpoint": ("microsoft powerpoint", "powerpoint"),
    "keynote": ("keynote",),
    "libreoffice": ("libreoffice",),
    "wps office": ("wps office", "wpsoffice", "wps"),
}


def now() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


def normalize_name(value: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", value.casefold())


def first_line(value: str) -> str | None:
    return next((line.strip() for line in value.splitlines() if line.strip()), None)


def atomic_write_json(path: Path, payload: dict[str, Any], compact: bool) -> None:
    path = path.expanduser().resolve()
    path.parent.mkdir(parents=True, exist_ok=True)
    indent = None if compact else 2
    text = json.dumps(payload, ensure_ascii=False, indent=indent)
    temp = path.with_suffix(path.suffix + ".tmp")
    temp.write_text(text + "\n", encoding="utf-8")
    os.replace(temp, path)


def probe_command(path: str | None, args: list[str], timeout: float) -> dict[str, Any]:
    result: dict[str, Any] = {"found": bool(path), "path": path, "usable": False}
    if not path:
        return result
    try:
        completed = subprocess.run(
            [path, *args],
            check=False,
            capture_output=True,
            text=True,
            timeout=timeout,
        )
        combined = "\n".join(part for part in (completed.stdout, completed.stderr) if part)
        result.update(
            {
                "returncode": completed.returncode,
                "usable": completed.returncode == 0,
                "summary": first_line(combined),
            }
        )
    except (OSError, subprocess.TimeoutExpired) as exc:
        result["error"] = str(exc)
    return result


def detect_editppt(timeout: float) -> dict[str, Any]:
    path = locate_tool("editppt")
    result = probe_command(path, ["--help"], timeout)
    result["id"] = "editppt"
    return result


def detect_libreoffice(timeout: float) -> dict[str, Any]:
    path = locate_tool("libreoffice")
    result = probe_command(path, ["--version"], timeout)
    result["id"] = "libreoffice"
    return result


def detect_pdftoppm(timeout: float) -> dict[str, Any]:
    path = locate_tool("pdftoppm")
    result = probe_command(path, ["-v"], timeout)
    # Poppler prints its version and may return either 0 or 99 depending on
    # packaging. A discovered executable with a version summary is usable.
    if path and result.get("summary") and result.get("returncode") in {0, 99}:
        result["usable"] = True
    result["id"] = "pdftoppm"
    return result


def detect_offline_ocr(timeout: float) -> dict[str, Any]:
    swiftc = locate_tool("swiftc") if platform.system() == "Darwin" else None
    script = Path(__file__).with_name("offline_ocr.swift")
    result: dict[str, Any] = {
        "id": "offline-ocr",
        "found": bool(swiftc and script.is_file()),
        "path": swiftc,
        "script": str(script.resolve()),
        "usable": False,
    }
    if not swiftc or not script.is_file():
        return result
    try:
        completed = subprocess.run(
            [swiftc, "-typecheck", str(script)],
            check=False,
            capture_output=True,
            text=True,
            timeout=max(timeout, 30.0),
        )
        result.update(
            {
                "returncode": completed.returncode,
                "usable": completed.returncode == 0,
                "summary": first_line("\n".join((completed.stdout, completed.stderr))),
                "backend": "apple-vision-offline",
                "external_service": False,
                "token_required": False,
            }
        )
    except (OSError, subprocess.TimeoutExpired) as exc:
        result["error"] = str(exc)
    return result


def inventory_fonts(timeout: float) -> tuple[list[dict[str, str]], list[str]]:
    records: dict[tuple[str, str], dict[str, str]] = {}
    methods: list[str] = []
    fc_list = shutil.which("fc-list")
    if fc_list:
        try:
            completed = subprocess.run(
                [fc_list, "--format", "%{family}\t%{file}\n"],
                check=False,
                capture_output=True,
                text=True,
                timeout=max(timeout, 8.0),
            )
            if completed.returncode == 0:
                methods.append("fontconfig")
                for line in completed.stdout.splitlines():
                    family, separator, raw_path = line.partition("\t")
                    if not separator:
                        continue
                    for family_name in (x.strip() for x in family.split(",")):
                        if family_name:
                            key = (normalize_name(family_name), raw_path)
                            records[key] = {"family": family_name, "path": raw_path}
        except (OSError, subprocess.TimeoutExpired):
            pass

    if not records:
        methods.append("filename-scan")
        for root in font_roots():
            if not root.is_dir():
                continue
            try:
                paths = root.rglob("*")
                for path in paths:
                    if path.is_file() and path.suffix.casefold() in FONT_EXTENSIONS:
                        family = path.stem
                        key = (normalize_name(family), str(path))
                        records[key] = {"family": family, "path": str(path)}
            except OSError:
                continue
    return sorted(records.values(), key=lambda item: (item["family"].casefold(), item["path"])), methods


def match_font(query: str, fonts: list[dict[str, str]]) -> list[dict[str, str]]:
    wanted = normalize_name(query)
    matches = []
    for item in fonts:
        family = normalize_name(item["family"])
        filename = normalize_name(Path(item["path"]).stem)
        if wanted and (wanted == family or wanted in family or wanted in filename):
            matches.append(item)
    return matches[:8]


def app_roots() -> list[Path]:
    return [Path("/Applications"), Path("/System/Applications"), Path.home() / "Applications"]


def inventory_apps() -> list[dict[str, str]]:
    apps: dict[str, dict[str, str]] = {}
    for root in app_roots():
        if not root.is_dir():
            continue
        try:
            for path in root.iterdir():
                if path.is_dir() and path.suffix.casefold() == ".app":
                    name = path.stem
                    apps[str(path.resolve())] = {"name": name, "path": str(path.resolve())}
        except OSError:
            continue
    return sorted(apps.values(), key=lambda item: item["name"].casefold())


def match_app(query: str, apps: list[dict[str, str]]) -> list[dict[str, str]]:
    direct = Path(query).expanduser()
    if direct.suffix.casefold() == ".app" and direct.is_dir():
        return [{"name": direct.stem, "path": str(direct.resolve())}]
    normalized = normalize_name(query)
    aliases = APP_ALIASES.get(query.casefold(), (query,))
    wanted = {normalized, *(normalize_name(alias) for alias in aliases)}
    exact = [item for item in apps if normalize_name(item["name"]) in wanted]
    if exact:
        return exact
    return [
        item
        for item in apps
        if any(token and token in normalize_name(item["name"]) for token in wanted)
    ]


def add_issue(container: list[str], message: str) -> None:
    if message not in container:
        container.append(message)


def check_status(errors: list[str], warnings: list[str]) -> str:
    if errors:
        return "failed"
    if warnings:
        return "passed_with_warnings"
    return "passed"


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Check ppt-gen environment, local applications, fonts, and privacy constraints.",
        epilog="Exit codes: 0 passed, 2 passed_with_warnings, 1 failed.",
    )
    parser.add_argument(
        "--data-handling",
        "--privacy",
        "--privacy-mode",
        dest="data_handling",
        choices=("standard", "local-only"),
        default="standard",
        help="Data-handling mode; local-only forbids remote image generation/uploads, while OCR always stays local/offline.",
    )
    parser.add_argument("--project-dir", type=Path, default=Path.cwd())
    parser.add_argument("--target", action="append", choices=STAGES, help="Requested workflow stage; repeat as needed.")
    parser.add_argument("--font", action="append", default=[], help="Required font family; repeat as needed.")
    parser.add_argument("--app", action="append", default=[], help="Optional app to detect; repeat as needed.")
    parser.add_argument("--require-app", action="append", default=[], help="Required app name or .app path.")
    parser.add_argument("--require-editppt", action="store_true", help="Fail unless editppt is callable.")
    parser.add_argument("--require-libreoffice", action="store_true", help="Fail unless LibreOffice is callable.")
    parser.add_argument(
        "--require-docx",
        action="store_true",
        help="Require LibreOffice plus pdftoppm, the local DOCX render/comparison path.",
    )
    parser.add_argument(
        "--require-pptx",
        action="store_true",
        help="Require LibreOffice plus pdftoppm, the local PPTX render/comparison path.",
    )
    parser.add_argument(
        "--require-editable",
        action="store_true",
        help="Require editppt plus LibreOffice/pdftoppm for editable-deck QA.",
    )
    parser.add_argument("--timeout", type=float, default=5.0, help="Seconds allowed for each command probe.")
    parser.add_argument("--json-out", type=Path, help="Also write the JSON result to this path.")
    parser.add_argument("--compact", action="store_true", help="Emit compact JSON.")
    return parser


def main() -> int:
    args = build_parser().parse_args()
    targets = list(dict.fromkeys(args.target or []))
    errors: list[str] = []
    warnings: list[str] = []
    checks: list[dict[str, Any]] = []

    project_dir = args.project_dir.expanduser().resolve()
    nearest = project_dir
    while not nearest.exists() and nearest != nearest.parent:
        nearest = nearest.parent
    writable = nearest.is_dir() and os.access(nearest, os.W_OK)
    project_check = {
        "id": "project-directory",
        "status": "passed" if writable else "failed",
        "path": str(project_dir),
        "nearest_existing_parent": str(nearest),
        "writable": writable,
    }
    checks.append(project_check)
    if not writable:
        add_issue(errors, f"Project directory is not writable: {project_dir}")

    pillow_found = importlib.util.find_spec("PIL") is not None
    pillow_required = bool(args.require_docx or args.require_pptx or args.require_editable or set(targets) & {"style-options", "template", "image-deck", "editable-deck"})
    checks.append({"id": "python-pillow", "required": pillow_required,
                   "found": pillow_found, "executable": sys.executable,
                   "status": "failed" if pillow_required and not pillow_found else "passed"})
    if pillow_required and not pillow_found:
        add_issue(errors, "Pillow is missing from this Python interpreter; use the bundled artifact runtime or an environment with Pillow installed.")

    privacy_details: dict[str, Any] = {
        "mode": args.data_handling,
        "remote_image_generation_allowed": args.data_handling == "standard",
        "external_ocr_allowed": False,
        "ocr_policy": "offline",
        "ask_for_ocr_token": False,
        "secret_values_recorded": False,
        "constraints": [],
    }
    privacy_warnings: list[str] = []
    if args.data_handling == "local-only":
        privacy_details["constraints"] = [
            "Do not call built-in or third-party image-generation services.",
            "Do not upload source pages to external OCR services (also prohibited in standard mode).",
            "Use local/native composition and offline text hints only.",
        ]
        if set(targets) & {"style-options", "template", "image-deck"}:
            privacy_warnings.append("Selected visual stages require a local/native visual route; built-in imagegen is disabled.")
        if "editable-deck" in targets:
            privacy_warnings.append("Editable conversion must disable every remote image backend and external OCR path.")
    for message in privacy_warnings:
        add_issue(warnings, message)
    checks.append(
        {
            "id": "privacy",
            "status": "passed_with_warnings" if privacy_warnings else "passed",
            **privacy_details,
            "warnings": privacy_warnings,
        }
    )

    editppt_required = args.require_editppt or args.require_editable or "editable-deck" in targets
    editppt = detect_editppt(args.timeout)
    editppt["required"] = editppt_required
    if not editppt["usable"]:
        message = "editppt is unavailable or its --help probe failed."
        if editppt_required:
            add_issue(errors, message)
            editppt["status"] = "failed"
        else:
            add_issue(warnings, message)
            editppt["status"] = "passed_with_warnings"
    else:
        editppt["status"] = "passed"
    checks.append(editppt)

    libreoffice_recommended = bool(
        set(targets)
        & {"report", "image-deck", "editable-deck", "speaker-notes", "qa-package", "deck-revision"}
    )
    libreoffice_required = (
        args.require_libreoffice or args.require_docx or args.require_pptx or args.require_editable
    )
    libreoffice = detect_libreoffice(args.timeout)
    libreoffice["required"] = libreoffice_required
    libreoffice["recommended"] = libreoffice_recommended
    if not libreoffice["usable"]:
        message = "LibreOffice/soffice is unavailable or its version probe failed."
        if libreoffice_required:
            add_issue(errors, message)
            libreoffice["status"] = "failed"
        elif libreoffice_recommended:
            add_issue(warnings, message + " Render-based QA will need an approved fallback.")
            libreoffice["status"] = "passed_with_warnings"
        else:
            libreoffice["status"] = "passed"
    else:
        libreoffice["status"] = "passed"
    checks.append(libreoffice)

    pdftoppm_required = bool(args.require_docx or args.require_pptx or args.require_editable)
    pdftoppm_recommended = libreoffice_recommended
    pdftoppm = detect_pdftoppm(args.timeout)
    pdftoppm["required"] = pdftoppm_required
    pdftoppm["recommended"] = pdftoppm_recommended
    if not pdftoppm["usable"]:
        message = "Poppler pdftoppm is unavailable or its version probe failed."
        if pdftoppm_required:
            add_issue(errors, message)
            pdftoppm["status"] = "failed"
        elif pdftoppm_recommended:
            add_issue(warnings, message + " Independent render comparison is unavailable.")
            pdftoppm["status"] = "passed_with_warnings"
        else:
            pdftoppm["status"] = "passed"
    else:
        pdftoppm["status"] = "passed"
    checks.append(pdftoppm)

    offline_ocr_required = bool(args.require_editable or set(targets) & {"image-deck", "editable-deck"})
    offline_ocr = detect_offline_ocr(args.timeout)
    offline_ocr["required"] = offline_ocr_required
    if not offline_ocr["usable"]:
        message = "Local Apple Vision OCR verifier is unavailable (requires macOS and Xcode Command Line Tools; no online fallback)."
        if offline_ocr_required:
            add_issue(errors, message)
            offline_ocr["status"] = "failed"
        else:
            offline_ocr["status"] = "passed"
    else:
        offline_ocr["status"] = "passed"
    checks.append(offline_ocr)

    fonts, font_methods = inventory_fonts(args.timeout)
    requested_fonts = []
    font_errors = []
    for name in args.font:
        matches = match_font(name, fonts)
        requested_fonts.append({"query": name, "found": bool(matches), "matches": matches})
        if not matches:
            message = f"Required font was not found: {name}"
            font_errors.append(message)
            add_issue(errors, message)
    font_status = "failed" if font_errors else ("passed" if fonts else "passed_with_warnings")
    if not fonts:
        add_issue(warnings, "No font inventory could be produced.")
    checks.append(
        {
            "id": "fonts",
            "status": font_status,
            "detectors": font_methods,
            "installed_family_records": len(fonts),
            "requested": requested_fonts,
            "sample_families": sorted({item["family"] for item in fonts}, key=str.casefold)[:20],
        }
    )

    apps = inventory_apps()
    optional_queries = list(dict.fromkeys([*DEFAULT_APPS, *args.app]))
    required_queries = list(dict.fromkeys(args.require_app))
    app_results = []
    for name in optional_queries:
        matches = match_app(name, apps)
        app_results.append({"query": name, "required": False, "found": bool(matches), "matches": matches})
    for name in required_queries:
        matches = match_app(name, apps)
        app_results.append({"query": name, "required": True, "found": bool(matches), "matches": matches})
        if not matches:
            add_issue(errors, f"Required application was not found: {name}")
    known_available = any(item["found"] for item in app_results if not item["required"])
    app_status = "failed" if any(item["required"] and not item["found"] for item in app_results) else "passed"
    if not known_available and not required_queries:
        add_issue(warnings, "No common presentation application was found in standard macOS app locations.")
        app_status = "passed_with_warnings"
    checks.append(
        {
            "id": "applications",
            "status": app_status,
            "installed_app_count": len(apps),
            "requested": app_results,
        }
    )

    status = check_status(errors, warnings)
    next_actions: list[str] = []
    if not editppt["usable"] and editppt_required:
        next_actions.append("Install or repair editppt before starting editable-deck reconstruction.")
    if not libreoffice["usable"] and (libreoffice_recommended or libreoffice_required):
        next_actions.append("Install LibreOffice or record that render-based DOCX/PPTX QA is unavailable.")
    if not pdftoppm["usable"] and (pdftoppm_recommended or pdftoppm_required):
        next_actions.append("Install Poppler/pdftoppm before final DOCX/PPTX render comparison.")
    if offline_ocr_required and not offline_ocr["usable"]:
        next_actions.append("Restore the local Swift/Apple Vision OCR verifier before image-slide final QA.")
    if args.data_handling == "local-only":
        next_actions.append("Propagate local-only to every companion and do not invoke remote image/OCR backends.")
    if font_errors:
        next_actions.append("Install the required fonts or select verified substitutes before layout work.")
    if set(targets) & {"style-options", "template", "image-deck", "editable-deck"}:
        if args.data_handling == "standard":
            next_actions.append(
                "Confirm the active agent runtime exposes built-in imagegen before visual generation; CLI preflight cannot probe agent-only tools."
            )
        if "editable-deck" in targets:
            next_actions.append(
                "Confirm multi-agent page dispatch is available before starting a multi-page image-to-editable conversion."
            )
    if set(targets) & {"report", "image-deck", "editable-deck", "speaker-notes", "qa-package", "deck-revision"}:
        next_actions.append(
            "Run a task-local Chinese DOCX/PPTX render smoke test before production; executable discovery alone does not verify glyph fidelity."
        )

    payload: dict[str, Any] = {
        "schema_version": 1,
        "tool": "ppt-gen-preflight",
        "generated_at": now(),
        "status": status,
        "data_handling": args.data_handling,
        "privacy_mode": args.data_handling,
        "requirements": {
            "docx": args.require_docx,
            "pptx": args.require_pptx,
            "editable": args.require_editable,
            "mapped_checks": {
                "docx": ["libreoffice", "pdftoppm"],
                "pptx": ["libreoffice", "pdftoppm"],
                "editable": ["editppt", "libreoffice", "pdftoppm"],
                "image-text-verification": ["offline-ocr"],
            },
        },
        "targets": targets,
        "environment": {
            "platform": platform.platform(),
            "python": sys.version.split()[0],
            "executable": sys.executable,
        },
        "checks": checks,
        "errors": errors,
        "warnings": warnings,
        "next_actions": next_actions,
    }
    output = json.dumps(payload, ensure_ascii=False, indent=None if args.compact else 2)
    print(output)
    if args.json_out:
        atomic_write_json(args.json_out, payload, args.compact)
    return {"passed": 0, "passed_with_warnings": 2, "failed": 1}[status]


if __name__ == "__main__":
    raise SystemExit(main())
