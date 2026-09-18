#!/usr/bin/env python3
"""Plan page-local raster reuse. Reuse candidates still require current final QA."""
from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path

from local_runtime import sha256


def copy_hash(page: dict) -> str:
    return hashlib.sha256(json.dumps(page, ensure_ascii=False, sort_keys=True, separators=(",", ":")).encode()).hexdigest()


def validate_pages(pages: list) -> None:
    if not isinstance(pages, list) or not pages:
        raise ValueError("slides must be a nonempty array")
    ids = []
    for index, page in enumerate(pages, 1):
        if page.get("page") != index or not isinstance(page.get("id"), str) or not page["id"]:
            raise ValueError("slides require contiguous page numbers and stable ids")
        ids.append(page["id"])
    if len(set(ids)) != len(ids):
        raise ValueError("slide ids must be unique")


def make_index(ledger: dict, visual: Path, manifest: dict, qa: Path) -> dict:
    validate_pages(ledger.get("slides"))
    report = json.loads(qa.read_text(encoding="utf-8"))
    if report.get("mode") != "final-delivery" or report.get("passed") is not True or report.get("errors"):
        raise ValueError("reuse index requires a passing non-structural QA report")
    image_entries = manifest.get("slides", [])
    if len(image_entries) != len(ledger["slides"]):
        raise ValueError("manifest and copy ledger page counts differ")
    pages = []
    for copy, image in zip(ledger["slides"], image_entries):
        if copy["page"] != image.get("page") or copy["id"] != image.get("id"):
            raise ValueError("manifest/copy page identity mismatch")
        path = Path(image["path"]).expanduser().resolve()
        if not path.is_file() or sha256(path) != image.get("sha256"):
            raise ValueError("approved page image is missing or changed")
        pages.append({"page": copy["page"], "id": copy["id"], "copy_sha256": copy_hash(copy),
                      "image": {"path": str(path), "sha256": sha256(path)}})
    return {"schema_version": 1, "visual_sha256": sha256(visual),
            "qa": {"path": str(qa.resolve()), "sha256": sha256(qa)}, "pages": pages}


def plan(ledger: dict, visual: Path, index: dict) -> dict:
    validate_pages(ledger.get("slides"))
    if index.get("schema_version") != 1:
        raise ValueError("unsupported reuse-index schema")
    old_pages = index.get("pages", [])
    if len({page["id"] for page in old_pages}) != len(old_pages):
        raise ValueError("reuse index contains duplicate ids")
    previous = {page["id"]: page for page in old_pages}
    qa = index.get("qa", {})
    qa_path = Path(qa.get("path", ""))
    prior_qa_intact = qa_path.is_file() and sha256(qa_path) == qa.get("sha256")
    visual_unchanged = sha256(visual) == index.get("visual_sha256")
    output = []
    for page in ledger["slides"]:
        old = previous.get(page["id"], {})
        image = old.get("image", {})
        image_path = Path(image.get("path", ""))
        reasons = []
        if not prior_qa_intact:
            reasons.append("prior QA missing/changed")
        if not visual_unchanged:
            reasons.append("visual system changed")
        if old.get("copy_sha256") != copy_hash(page):
            reasons.append("page copy/identity changed")
        if not image_path.is_file() or sha256(image_path) != image.get("sha256"):
            reasons.append("approved image missing/changed")
        output.append({"page": page["page"], "id": page["id"],
                       "action": "regenerate" if reasons else "reuse-candidate",
                       "reasons": reasons, "image": image if not reasons else None})
    return {"schema_version": 1, "requires_current_final_qa": True, "pages": output,
            "removed_ids": sorted(set(previous) - {page["id"] for page in ledger["slides"]})}


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("mode", choices=("record", "plan"))
    parser.add_argument("--ledger", type=Path, required=True)
    parser.add_argument("--visual-system", type=Path, required=True)
    parser.add_argument("--index", type=Path, required=True)
    parser.add_argument("--slide-manifest", type=Path)
    parser.add_argument("--qa-report", type=Path)
    args = parser.parse_args()
    ledger = json.loads(args.ledger.read_text(encoding="utf-8"))
    if args.mode == "record":
        if not args.slide_manifest or not args.qa_report:
            parser.error("record requires --slide-manifest and --qa-report")
        qa = json.loads(args.qa_report.read_text(encoding="utf-8"))
        for role, path in (("slide-copy-ledger", args.ledger), ("slide-manifest", args.slide_manifest), ("visual-system", args.visual_system)):
            if not any(item.get("role") == role and item.get("path") == str(path.resolve()) and item.get("sha256") == sha256(path) for item in qa.get("artifacts", [])):
                parser.error(f"QA does not bind the current {role}")
        output = make_index(ledger, args.visual_system, json.loads(args.slide_manifest.read_text(encoding="utf-8")), args.qa_report)
        if args.index.exists():
            parser.error("reuse index exists; write a new version to preserve the previous snapshot")
        args.index.parent.mkdir(parents=True, exist_ok=True)
        args.index.write_text(json.dumps(output, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    else:
        output = plan(ledger, args.visual_system, json.loads(args.index.read_text(encoding="utf-8")))
    print(json.dumps(output, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
