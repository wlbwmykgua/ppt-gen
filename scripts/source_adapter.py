#!/usr/bin/env python3
"""Inspect role=path sources and initialize ppt-gen authority manifests.

The adapter identifies content from file signatures/package contents, hashes every
source, writes ``source-manifest.json``, and creates JSON ledgers only when they do
not already exist (unless ``--force-ledgers`` is explicitly supplied).
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import posixpath
import re
import struct
import zipfile
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable
from xml.etree import ElementTree as ET


IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".webp", ".tif", ".tiff", ".bmp"}
GOALS = ("auto", "report", "slides", "editable-deck", "speaker-notes", "revision")
POLICIES = ("auto", "faithful", "verify-update")
SLIDE_RE = re.compile(r"ppt/slides/slide(\d+)\.xml$")
NUMBER_RE = re.compile(r"(\d+)")


def now() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


def local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def natural_key(value: str) -> list[Any]:
    return [int(part) if part.isdigit() else part.casefold() for part in NUMBER_RE.split(value)]


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def hash_directory(path: Path) -> tuple[str, int, int]:
    digest = hashlib.sha256()
    total_size = 0
    count = 0
    entries = list(path.rglob("*"))
    symlinks = [entry for entry in entries if entry.is_symlink()]
    if symlinks:
        relative = ", ".join(sorted(entry.relative_to(path).as_posix() for entry in symlinks))
        raise ValueError(f"source directories may not contain symlinks: {relative}")
    for file_path in sorted((p for p in entries if p.is_file()), key=lambda p: natural_key(p.relative_to(path).as_posix())):
        relative = file_path.relative_to(path).as_posix()
        file_hash = sha256_file(file_path)
        size = file_path.stat().st_size
        digest.update(relative.encode("utf-8"))
        digest.update(b"\0")
        digest.update(str(size).encode("ascii"))
        digest.update(b"\0")
        digest.update(file_hash.encode("ascii"))
        digest.update(b"\n")
        total_size += size
        count += 1
    return digest.hexdigest(), total_size, count


def atomic_write_json(path: Path, payload: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_suffix(path.suffix + ".tmp")
    temp.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    os.replace(temp, path)


def xml_text(raw: bytes) -> list[str]:
    try:
        root = ET.fromstring(raw)
    except ET.ParseError:
        return []
    return [
        (node.text or "").strip()
        for node in root.iter()
        if local_name(node.tag) == "t" and (node.text or "").strip()
    ]


def png_size(path: Path) -> tuple[int, int]:
    data = path.read_bytes()[:24]
    if len(data) < 24 or data[:8] != b"\x89PNG\r\n\x1a\n":
        raise ValueError("not a PNG")
    return struct.unpack(">II", data[16:24])


def jpeg_size(path: Path) -> tuple[int, int]:
    with path.open("rb") as handle:
        if handle.read(2) != b"\xff\xd8":
            raise ValueError("not a JPEG")
        while True:
            marker = handle.read(1)
            if not marker:
                break
            if marker != b"\xff":
                continue
            code = handle.read(1)
            while code == b"\xff":
                code = handle.read(1)
            if not code:
                break
            value = code[0]
            if value in {0xD8, 0xD9}:
                continue
            length_bytes = handle.read(2)
            if len(length_bytes) != 2:
                break
            length = struct.unpack(">H", length_bytes)[0]
            payload = handle.read(length - 2)
            if value in {0xC0, 0xC1, 0xC2, 0xC3, 0xC5, 0xC6, 0xC7, 0xC9, 0xCA, 0xCB, 0xCD, 0xCE, 0xCF}:
                height, width = struct.unpack(">HH", payload[1:5])
                return width, height
    raise ValueError("JPEG dimensions not found")


def image_size(path: Path) -> tuple[int, int]:
    try:
        from PIL import Image

        with Image.open(path) as image:
            image.verify()
        with Image.open(path) as image:
            image.load()
            return int(image.width), int(image.height)
    except ImportError:
        if path.suffix.casefold() == ".png":
            return png_size(path)
        if path.suffix.casefold() in {".jpg", ".jpeg"}:
            return jpeg_size(path)
        raise


def image_record(path: Path, base: Path | None = None) -> dict[str, Any]:
    width, height = image_size(path)
    record: dict[str, Any] = {
        "path": str(path.resolve()),
        "sha256": sha256_file(path),
        "size": path.stat().st_size,
        "width": width,
        "height": height,
        "exact_16_9": width * 9 == height * 16,
    }
    if base:
        record["relative_path"] = path.relative_to(base).as_posix()
    return record


def inspect_docx(archive: zipfile.ZipFile) -> dict[str, Any]:
    raw = archive.read("word/document.xml")
    root = ET.fromstring(raw)
    paragraphs: list[str] = []
    headings: list[dict[str, Any]] = []
    for paragraph in (node for node in root.iter() if local_name(node.tag) == "p"):
        text = "".join((node.text or "") for node in paragraph.iter() if local_name(node.tag) == "t").strip()
        if not text:
            continue
        paragraphs.append(text)
        style = next((node for node in paragraph.iter() if local_name(node.tag) == "pStyle"), None)
        style_value = ""
        if style is not None:
            style_value = next((value for key, value in style.attrib.items() if local_name(key) == "val"), "")
        if style_value.casefold().startswith("heading") or style_value.startswith("标题"):
            level_match = NUMBER_RE.search(style_value)
            headings.append({"level": int(level_match.group(1)) if level_match else 1, "text": text})
    text = "\n".join(paragraphs)
    return {
        "kind": "docx",
        "paragraph_count": len(paragraphs),
        "heading_count": len(headings),
        "headings": headings[:100],
        "text_characters": len(text),
        "text_preview": text[:1200],
        "_text": text,
    }


def presentation_slide_paths(archive: zipfile.ZipFile) -> list[str]:
    names = set(archive.namelist())
    try:
        presentation = ET.fromstring(archive.read("ppt/presentation.xml"))
        relationships = ET.fromstring(archive.read("ppt/_rels/presentation.xml.rels"))
        rel_map = {
            node.attrib.get("Id"): node.attrib.get("Target")
            for node in relationships
            if node.attrib.get("Id") and node.attrib.get("Target")
        }
        ordered = []
        for node in presentation.iter():
            if local_name(node.tag) != "sldId":
                continue
            rel_id = next(
                (
                    value
                    for key, value in node.attrib.items()
                    if local_name(key) == "id" and (key.startswith("{") or ":" in key)
                ),
                None,
            )
            target = rel_map.get(rel_id or "")
            if target:
                normalized = (
                    posixpath.normpath(target.lstrip("/"))
                    if target.startswith("/")
                    else posixpath.normpath(posixpath.join("ppt", target))
                )
                if normalized in names:
                    ordered.append(normalized)
        if ordered:
            return ordered
    except (KeyError, ET.ParseError):
        pass
    return sorted((name for name in names if SLIDE_RE.match(name)), key=lambda name: int(SLIDE_RE.match(name).group(1)))


def notes_for_slide(archive: zipfile.ZipFile, slide_path: str) -> str:
    slide_name = posixpath.basename(slide_path)
    rel_path = f"ppt/slides/_rels/{slide_name}.rels"
    if rel_path not in archive.namelist():
        return ""
    try:
        root = ET.fromstring(archive.read(rel_path))
    except ET.ParseError:
        return ""
    for relationship in root:
        if not relationship.attrib.get("Type", "").endswith("/notesSlide"):
            continue
        target = relationship.attrib.get("Target")
        if not target:
            continue
        notes_path = (
            posixpath.normpath(target.lstrip("/"))
            if target.startswith("/")
            else posixpath.normpath(posixpath.join(posixpath.dirname(slide_path), target))
        )
        if notes_path in archive.namelist():
            return "\n".join(xml_text(archive.read(notes_path))).strip()
    return ""


def inspect_pptx(archive: zipfile.ZipFile) -> dict[str, Any]:
    slide_paths = presentation_slide_paths(archive)
    slides: list[dict[str, Any]] = []
    image_only_count = 0
    notes_count = 0
    for index, slide_path in enumerate(slide_paths, start=1):
        raw = archive.read(slide_path)
        root = ET.fromstring(raw)
        tags = [local_name(node.tag) for node in root.iter()]
        texts = xml_text(raw)
        picture_count = tags.count("pic")
        editable_count = tags.count("sp") + tags.count("graphicFrame") + tags.count("cxnSp")
        image_only = picture_count == 1 and editable_count == 0
        if image_only:
            image_only_count += 1
        notes = notes_for_slide(archive, slide_path)
        if notes:
            notes_count += 1
        slides.append(
            {
                "page_number": index,
                "source_part": slide_path,
                "title": texts[0] if texts else "",
                "visible_text": texts,
                "notes": notes,
                "picture_count": picture_count,
                "editable_object_count": editable_count,
                "image_only": image_only,
            }
        )
    wholly_image_only = bool(slides) and image_only_count == len(slides)
    return {
        "kind": "pptx-image-only" if wholly_image_only else "pptx-editable",
        "slide_count": len(slides),
        "image_only_slide_count": image_only_count,
        "notes_slide_count": notes_count,
        "text_characters": sum(len(text) for slide in slides for text in slide["visible_text"]),
        "_slides": slides,
    }


def inspect_zip_office(path: Path) -> dict[str, Any] | None:
    try:
        with zipfile.ZipFile(path) as archive:
            names = set(archive.namelist())
            if "word/document.xml" in names:
                return inspect_docx(archive)
            if "ppt/presentation.xml" in names:
                return inspect_pptx(archive)
    except (OSError, zipfile.BadZipFile, ET.ParseError, KeyError, ValueError):
        return None
    return None


def inspect_pdf(path: Path) -> dict[str, Any]:
    page_count: int | None = None
    text = ""
    detector = "regex"
    try:
        import fitz

        document = fitz.open(str(path))
        try:
            page_count = document.page_count
            text = "\n".join(document.load_page(index).get_text("text") for index in range(min(page_count, 50)))
        finally:
            document.close()
        detector = "pymupdf"
    except (ImportError, OSError, RuntimeError, ValueError):
        try:
            from pypdf import PdfReader

            reader = PdfReader(str(path))
            page_count = len(reader.pages)
            pieces = []
            for page in reader.pages[:50]:
                try:
                    pieces.append(page.extract_text() or "")
                except Exception:
                    pieces.append("")
            text = "\n".join(pieces)
            detector = "pypdf"
        except (ImportError, OSError, ValueError):
            raw = path.read_bytes()
            page_count = len(re.findall(rb"/Type\s*/Page(?!s)\b", raw)) or None
    return {
        "kind": "pdf",
        "page_count": page_count,
        "text_characters": len(text),
        "text_preview": text[:1200],
        "text_detector": detector,
        "_text": text,
    }


def inspect_markdown(path: Path) -> dict[str, Any]:
    text = path.read_text(encoding="utf-8", errors="replace")
    headings = []
    for line in text.splitlines():
        match = re.match(r"^(#{1,6})\s+(.+?)\s*$", line)
        if match:
            headings.append({"level": len(match.group(1)), "text": match.group(2).strip()})
    return {
        "kind": "markdown",
        "line_count": len(text.splitlines()),
        "heading_count": len(headings),
        "headings": headings[:100],
        "text_characters": len(text),
        "text_preview": text[:1200],
        "_text": text,
    }


def inspect_file(path: Path) -> dict[str, Any]:
    try:
        header = path.read_bytes()[:16]
    except OSError as exc:
        return {"kind": "unknown", "inspection_error": str(exc)}
    if header.startswith(b"PK\x03\x04"):
        office = inspect_zip_office(path)
        if office:
            return office
        if path.suffix.casefold() in {".docx", ".pptx"}:
            return {"kind": "corrupt", "inspection_error": "OOXML package is unreadable or lacks its required document part."}
    if header.startswith(b"%PDF-"):
        return inspect_pdf(path)
    if path.suffix.casefold() in IMAGE_EXTENSIONS or header.startswith((b"\x89PNG", b"\xff\xd8")):
        try:
            width, height = image_size(path)
            return {
                "kind": "image",
                "width": width,
                "height": height,
                "exact_16_9": width * 9 == height * 16,
                "_images": [image_record(path)],
            }
        except (OSError, ValueError, struct.error) as exc:
            return {"kind": "corrupt", "inspection_error": f"Image inspection failed: {exc}"}
    if path.suffix.casefold() in {".md", ".markdown", ".mdown", ".txt"}:
        return inspect_markdown(path)
    try:
        sample = path.read_text(encoding="utf-8")
        if re.search(r"(?m)^#{1,6}\s+\S", sample) or re.search(r"(?m)^[-*+]\s+\S", sample):
            return inspect_markdown(path)
    except (OSError, UnicodeDecodeError):
        pass
    if path.suffix.casefold() in {".docx", ".pptx", ".pdf"} | IMAGE_EXTENSIONS:
        return {
            "kind": "corrupt",
            "inspection_error": f"Content signature does not match the declared {path.suffix.casefold() or 'file'} type.",
        }
    return {"kind": "unknown"}


def inspect_directory(path: Path) -> dict[str, Any]:
    images = sorted(
        (file_path for file_path in path.rglob("*") if file_path.is_file() and file_path.suffix.casefold() in IMAGE_EXTENSIONS),
        key=lambda file_path: natural_key(file_path.relative_to(path).as_posix()),
    )
    records = []
    warnings = []
    for image_path in images:
        try:
            records.append(image_record(image_path, path))
        except (OSError, ValueError, struct.error) as exc:
            warnings.append(f"Cannot inspect image {image_path.name}: {exc}")
    return {
        "kind": "images" if records else "directory",
        "image_count": len(records),
        "all_exact_16_9": bool(records) and all(item["exact_16_9"] for item in records),
        "warnings": warnings,
        "_images": records,
    }


def semantic_role(role: str, kind: str) -> str:
    normalized = role.casefold()
    if any(token in normalized for token in ("outline", "大纲")):
        return "outline"
    if any(token in normalized for token in ("template", "brand", "style", "模板", "品牌", "风格")):
        return "template"
    if any(token in normalized for token in ("report", "research", "报告", "研究")):
        return "report"
    if any(token in normalized for token in ("scan", "image", "screenshot", "扫描", "图片", "截图")):
        return "visual-slides"
    if any(token in normalized for token in ("deck", "ppt", "slides", "演示", "幻灯")):
        return "deck"
    if kind in {"images", "image", "pptx-image-only"}:
        return "visual-slides"
    if kind in {"docx", "markdown", "pdf"}:
        return "document"
    if kind == "pptx-editable":
        return "deck"
    return "other"


def media_type(kind: str) -> str:
    return {
        "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "pptx-editable": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        "pptx-image-only": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        "pdf": "application/pdf",
        "markdown": "text/markdown",
        "image": "image/*",
        "images": "application/x-image-sequence",
    }.get(kind, "application/octet-stream")


def suggested_start(record: dict[str, Any]) -> str:
    role = record.get("semantic_role")
    kind = record.get("kind")
    if role == "outline":
        return "style-options"
    if role == "template":
        return "template"
    if kind in {"image", "images", "pptx-image-only"} or role == "visual-slides":
        return "editable-deck"
    if kind in {"docx", "markdown", "pdf"} or role == "report":
        return "outline"
    if kind == "pptx-editable":
        return "deck-revision"
    return "brief"


def public_record(record: dict[str, Any]) -> dict[str, Any]:
    public = {key: value for key, value in record.items() if not key.startswith("_")}
    properties = {
        key: public[key]
        for key in (
            "page_count",
            "slide_count",
            "image_count",
            "paragraph_count",
            "heading_count",
            "text_characters",
            "all_exact_16_9",
            "exact_16_9",
        )
        if key in public
    }
    public.update(
        {
            "id": f"src-{record['_input_index'] + 1:03d}",
            "authority_rank": record["_input_index"] + 1,
            "media_type": media_type(str(record.get("kind"))),
            "adapter": record.get("kind"),
            "suggested_start": suggested_start(record),
            "properties": properties,
        }
    )
    return public


def inspect_input(role: str, raw_path: str) -> dict[str, Any]:
    path = Path(raw_path).expanduser().resolve()
    record: dict[str, Any] = {
        "role": role,
        "path": str(path),
        "exists": path.exists(),
    }
    if not path.exists():
        record.update({"kind": "missing", "semantic_role": "other"})
        return record
    stat = path.stat()
    record["modified_at"] = datetime.fromtimestamp(stat.st_mtime).astimezone().isoformat(timespec="seconds")
    if path.is_dir():
        digest, total_size, file_count = hash_directory(path)
        record.update({"sha256": digest, "size": total_size, "file_count": file_count})
        record.update(inspect_directory(path))
    elif path.is_file():
        record.update({"sha256": sha256_file(path), "size": stat.st_size})
        record.update(inspect_file(path))
    else:
        record["kind"] = "unsupported"
    record["semantic_role"] = semantic_role(role, record["kind"])
    return record


def parse_inputs(values: Iterable[str]) -> list[tuple[str, str]]:
    parsed = []
    for value in values:
        if "=" not in value:
            raise SystemExit(f"Input must use role=path: {value}")
        role, raw_path = value.split("=", 1)
        role = role.strip()
        raw_path = raw_path.strip()
        if not role or not raw_path:
            raise SystemExit(f"Input must use non-empty role=path: {value}")
        parsed.append((role, raw_path))
    if not parsed:
        raise SystemExit("Provide at least one role=path input")
    return parsed


def recommend_route(records: list[dict[str, Any]], goal: str) -> dict[str, Any]:
    existing = [record for record in records if record.get("exists")]
    kinds = {record.get("kind") for record in existing}
    roles = {record.get("semantic_role") for record in existing}
    if goal == "report":
        return {"entry_phase": "report", "confidence": "high", "reason": "The requested goal is a report."}
    if goal == "speaker-notes":
        return {"entry_phase": "speaker-notes", "confidence": "high", "reason": "The requested goal is speaker notes."}
    if goal == "editable-deck":
        if kinds & {"images", "image", "pptx-image-only"} or any(
            record.get("kind") == "pdf" and record.get("text_characters", 0) < 100 for record in existing
        ):
            return {"entry_phase": "editable-deck", "confidence": "high", "reason": "Visual slide sources are ready for reconstruction."}
        return {"entry_phase": "outline", "confidence": "medium", "reason": "Content sources must be turned into a slide story before native deck authoring."}
    if goal == "revision":
        return {
            "entry_phase": "deck-revision",
            "confidence": "high",
            "reason": "The requested goal is a direct revision of an existing deck.",
        }
    if "outline" in roles:
        return {"entry_phase": "style-options", "confidence": "high", "reason": "A page-by-page outline was supplied."}
    if "template" in roles and kinds & {"pptx-editable", "pptx-image-only", "image", "images"}:
        return {"entry_phase": "template", "confidence": "high", "reason": "A fixed visual authority was supplied."}
    if kinds & {"images", "image", "pptx-image-only"} or "visual-slides" in roles:
        return {"entry_phase": "editable-deck", "confidence": "high", "reason": "The strongest supplied authority is a visual slide source."}
    if "report" in roles or kinds & {"docx", "markdown"}:
        return {"entry_phase": "outline", "confidence": "high", "reason": "A substantive document can act as the slide-story authority."}
    text_pdf = any(record.get("kind") == "pdf" and record.get("text_characters", 0) >= 100 for record in existing)
    if text_pdf:
        return {"entry_phase": "outline", "confidence": "medium", "reason": "The PDF contains extractable document text."}
    scanned_pdf = any(record.get("kind") == "pdf" and record.get("text_characters", 0) < 100 for record in existing)
    if scanned_pdf:
        return {"entry_phase": "editable-deck", "confidence": "medium", "reason": "The PDF appears scanned or image-led; verify before reconstruction."}
    if "pptx-editable" in kinds:
        return {
            "entry_phase": "speaker-notes",
            "confidence": "low",
            "reason": "An editable PPTX is present; confirm whether the user wants notes or a direct revision.",
            "alternatives": ["deck-revision"],
        }
    if goal == "slides":
        return {"entry_phase": "report", "confidence": "low", "reason": "No substantive slide authority was detected."}
    return {"entry_phase": "brief", "confidence": "low", "reason": "The supplied sources do not establish a later entry phase."}


def recommend_policy(records: list[dict[str, Any]], requested: str, goal: str) -> dict[str, str]:
    if requested != "auto":
        return {"policy": requested, "reason": "Explicit command-line selection."}
    faithful_roles = {"outline", "template", "report", "visual-slides", "deck"}
    if goal in {"editable-deck", "speaker-notes", "revision"}:
        return {"policy": "faithful", "reason": "The goal operates on an existing content or visual authority."}
    if any(record.get("semantic_role") in faithful_roles for record in records if record.get("exists")):
        return {"policy": "faithful", "reason": "A report, outline, deck, template, or visual source should be preserved unless the user asks for updates."}
    return {"policy": "verify-update", "reason": "No fixed downstream authority was detected; verification and updating are appropriate."}


def authority_rank(record: dict[str, Any]) -> tuple[int, int]:
    role = record["role"].casefold()
    if any(token in role for token in ("primary", "authoritative", "主文件", "权威")):
        return (0, record["_input_index"])
    role_rank = {"report": 1, "outline": 2, "template": 3, "deck": 4, "visual-slides": 5}.get(record.get("semantic_role"), 9)
    return (role_rank, record["_input_index"])


def authority_order(records: list[dict[str, Any]]) -> list[str]:
    return [record["role"] for record in sorted(records, key=authority_rank)]


def primary_source_id(records: list[dict[str, Any]]) -> str | None:
    if not records:
        return None
    record = min(records, key=authority_rank)
    return f"src-{record['_input_index'] + 1:03d}"


def seed_slides(records: list[dict[str, Any]]) -> list[dict[str, Any]]:
    for record in records:
        slides = record.get("_slides")
        if slides:
            return [
                {
                    "page": slide["page_number"],
                    "page_number": slide["page_number"],
                    "id": f"slide-{slide['page_number']:03d}",
                    "title": slide.get("title", ""),
                    "purpose": "",
                    "body_copy": slide.get("visible_text", [])[1:],
                    "critical_fields": [],
                    "visual_type": "faithful reconstruction" if record["kind"] == "pptx-image-only" else "preserve existing slide",
                    "data": None,
                    "source_ids": [f"src-{record['_input_index'] + 1:03d}"],
                    "footer": "",
                    "chart_or_table_data": [],
                    "source_footer": "",
                    "visual_intent": "faithful reconstruction" if record["kind"] == "pptx-image-only" else "preserve existing slide",
                    "speaker_takeaway": "",
                    "claim_ids": [],
                    "claim_not_applicable_reason": "Draft adapter seed; classify and map claims before final QA.",
                    "source_role": record["role"],
                    "source_path": record["path"],
                }
                for slide in slides
            ]
    for record in records:
        images = record.get("_images")
        if images:
            return [
                {
                    "page": index,
                    "page_number": index,
                    "id": f"slide-{index:03d}",
                    "title": "",
                    "purpose": "",
                    "body_copy": [],
                    "critical_fields": [],
                    "visual_type": "faithful reconstruction",
                    "data": None,
                    "source_ids": [f"src-{record['_input_index'] + 1:03d}"],
                    "footer": "",
                    "chart_or_table_data": [],
                    "source_footer": "",
                    "visual_intent": "faithful reconstruction",
                    "speaker_takeaway": "",
                    "claim_ids": [],
                    "claim_not_applicable_reason": "Draft adapter seed; classify and map claims before final QA.",
                    "source_role": record["role"],
                    "source_image": image["path"],
                    "source_dimensions": [image["width"], image["height"]],
                }
                for index, image in enumerate(images, start=1)
            ]
    for record in records:
        if record.get("semantic_role") == "outline" and record.get("headings"):
            headings = [item for item in record["headings"] if item.get("level", 1) <= 2] or record["headings"]
            return [
                {
                    "page": index,
                    "page_number": index,
                    "id": f"slide-{index:03d}",
                    "title": heading["text"],
                    "purpose": "",
                    "body_copy": [],
                    "critical_fields": [],
                    "visual_type": "",
                    "data": None,
                    "source_ids": [f"src-{record['_input_index'] + 1:03d}"],
                    "footer": "",
                    "chart_or_table_data": [],
                    "source_footer": "",
                    "visual_intent": "",
                    "speaker_takeaway": "",
                    "claim_ids": [],
                    "claim_not_applicable_reason": "Draft adapter seed; classify and map claims before final QA.",
                    "source_role": record["role"],
                    "source_path": record["path"],
                }
                for index, heading in enumerate(headings, start=1)
            ]
    for record in records:
        page_count = record.get("page_count")
        if isinstance(page_count, int) and page_count > 0:
            return [
                {
                    "page": index,
                    "page_number": index,
                    "id": f"slide-{index:03d}",
                    "title": "",
                    "purpose": "",
                    "body_copy": [],
                    "critical_fields": [],
                    "visual_type": "",
                    "data": None,
                    "source_ids": [f"src-{record['_input_index'] + 1:03d}"],
                    "footer": "",
                    "chart_or_table_data": [],
                    "source_footer": "",
                    "visual_intent": "",
                    "speaker_takeaway": "",
                    "claim_ids": [],
                    "claim_not_applicable_reason": "Draft adapter seed; classify and map claims before final QA.",
                    "source_role": record["role"],
                    "source_path": record["path"],
                }
                for index in range(1, page_count + 1)
            ]
    return []


def seed_notes(records: list[dict[str, Any]], slides: list[dict[str, Any]]) -> list[dict[str, Any]]:
    note_map: dict[int, str] = {}
    for record in records:
        for slide in record.get("_slides", []):
            if slide.get("notes"):
                note_map[slide["page_number"]] = slide["notes"]
    seeded = []
    for slide in slides:
        script = note_map.get(slide["page_number"], "")
        target_seconds = parse_target_seconds(script)
        seeded.append(
            {
                "page": slide["page_number"],
                "page_number": slide["page_number"],
                "title": slide.get("title", ""),
                "target_seconds": target_seconds,
                "expected_seconds": target_seconds,
                "goal": "",
                "speaking_goal": "",
                "script": script,
                "cue": "",
                "delivery_cue": "",
                "transition": "",
                "claim_ids": [],
                "claim_not_applicable_reason": "Draft adapter seed; classify and map claims before final QA.",
                "source_caution": "",
            }
        )
    return seeded


def parse_target_seconds(script: str) -> int | None:
    if not script:
        return None
    label = r"(?:建议时长|预计时长|目标时长|时长|时间|target\s*time|duration)"
    clock = re.search(label + r"[^0-9]{0,12}(\d{1,3}):(\d{2})", script, flags=re.IGNORECASE)
    if clock:
        minutes, seconds = map(int, clock.groups())
        if seconds < 60:
            return minutes * 60 + seconds
    words = re.search(
        label + r"[^0-9]{0,12}(?:(\d{1,3})\s*分(?:钟)?)?\s*(?:(\d{1,3})\s*秒)?",
        script,
        flags=re.IGNORECASE,
    )
    if words and (words.group(1) or words.group(2)):
        return int(words.group(1) or 0) * 60 + int(words.group(2) or 0)
    return None


def initialize_ledgers(
    project_dir: Path,
    records: list[dict[str, Any]],
    policy: dict[str, str],
    data_handling: str,
    force: bool,
) -> dict[str, Any]:
    source_refs = [
        {
            "role": record["role"],
            "path": record["path"],
            "kind": record["kind"],
            "sha256": record.get("sha256"),
        }
        for record in records
        if record.get("exists")
    ]
    slides = seed_slides(records)
    notes = seed_notes(records, slides)
    notes_with_scripts = [note for note in notes if note["script"]]
    untimed_notes = [note["page"] for note in notes_with_scripts if note["target_seconds"] is None]
    note_warnings = (
        [
            "Imported speaker notes have no parseable target_seconds on pages "
            + ", ".join(map(str, untimed_notes))
            + "; those pages contribute 0 to total_seconds."
        ]
        if untimed_notes
        else []
    )
    payloads = {
        "claim-ledger.json": {
            "schema_version": 1,
            "status": "draft",
            "generated_at": now(),
            "source_cutoff": now()[:10],
            "source_manifest": str(project_dir / "source-manifest.json"),
            "content_policy": policy["policy"],
            "source_policy": policy["policy"],
            "data_handling": data_handling,
            "sources": source_refs,
            "claims": [],
            "not_applicable_reason": "Draft adapter seed; extract and verify claims before factual delivery.",
        },
        "slide-copy-ledger.json": {
            "schema_version": 1,
            "status": "draft",
            "generated_at": now(),
            "source_manifest": str(project_dir / "source-manifest.json"),
            "content_policy": policy["policy"],
            "source_policy": policy["policy"],
            "data_handling": data_handling,
            "deck": {"title": "", "canvas": "16:9", "source_cutoff": None},
            "slides": slides,
        },
        "speaker-notes.json": {
            "schema_version": 1,
            "status": "draft",
            "generated_at": now(),
            "source_manifest": str(project_dir / "source-manifest.json"),
            "content_policy": policy["policy"],
            "source_policy": policy["policy"],
            "data_handling": data_handling,
            "total_seconds": sum(note["target_seconds"] or 0 for note in notes),
            "warnings": note_warnings,
            "slides": notes,
        },
    }
    outcomes: dict[str, Any] = {}
    for filename, payload in payloads.items():
        path = project_dir / filename
        existed_before = path.exists()
        if path.exists() and not force:
            outcomes[filename] = {"status": "preserved", "path": str(path)}
            continue
        atomic_write_json(path, payload)
        outcomes[filename] = {
            "status": "replaced" if existed_before and force else "created",
            "path": str(path),
            "records": len(payload.get("claims", payload.get("slides", []))),
        }
    return outcomes


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Hash and classify role=path sources, then initialize ppt-gen JSON ledgers.",
        epilog="Example: source_adapter.py ./project report=brief.docx slides=images/ --goal slides",
    )
    parser.add_argument("project_dir", type=Path, help="Directory that will receive source-manifest.json and ledgers.")
    parser.add_argument("sources", nargs="*", metavar="role=path")
    parser.add_argument("--input", action="append", default=[], metavar="role=path", help="Additional source; repeat as needed.")
    parser.add_argument("--goal", choices=GOALS, default="auto")
    parser.add_argument(
        "--source-policy",
        "--policy",
        dest="source_policy",
        choices=POLICIES,
        default="auto",
        help="Content handling: faithful or verify-update.",
    )
    parser.add_argument(
        "--data-handling",
        choices=("standard", "local-only"),
        default="standard",
        help="Control allowed remote visual backends/uploads; OCR always remains offline and no-token.",
    )
    parser.add_argument("--skip-ledgers", action="store_true", help="Write only source-manifest.json.")
    parser.add_argument("--force-ledgers", action="store_true", help="Replace existing JSON ledgers.")
    return parser


def main() -> int:
    parser = build_parser()
    args, extras = parser.parse_known_args()
    unexpected = [value for value in extras if "=" not in value]
    if unexpected:
        parser.error(f"unrecognized arguments: {' '.join(unexpected)}")
    args.sources.extend(extras)
    pairs = parse_inputs([*args.sources, *args.input])
    project_dir = args.project_dir.expanduser().resolve()
    project_dir.mkdir(parents=True, exist_ok=True)

    records = []
    inspection_errors: list[str] = []
    for index, (role, raw_path) in enumerate(pairs):
        try:
            record = inspect_input(role, raw_path)
        except (OSError, ValueError) as exc:
            path = Path(raw_path).expanduser().resolve()
            record = {
                "role": role,
                "path": str(path),
                "exists": path.exists(),
                "kind": "corrupt",
                "semantic_role": "other",
                "inspection_error": str(exc),
            }
            inspection_errors.append(f"Source inspection failed: {role}={path} ({exc})")
        record["_input_index"] = index
        records.append(record)

    errors = [f"Source is missing: {record['role']}={record['path']}" for record in records if not record.get("exists")]
    errors.extend(inspection_errors)
    errors.extend(
        f"Source is corrupt: {record['role']}={record['path']} ({record.get('inspection_error', 'parsing failed')})"
        for record in records
        if record.get("kind") == "corrupt"
    )
    warnings: list[str] = []
    for record in records:
        warnings.extend(record.get("warnings", []))
        if record.get("kind") in {"unknown", "directory", "unsupported"}:
            warnings.append(f"Source type is not confidently recognized: {record['role']}={record['path']}")
        if record.get("kind") == "pdf" and record.get("text_detector") == "regex":
            warnings.append(f"PDF text extraction was unavailable for {record['role']}; scanned-versus-text routing is provisional.")
        untimed_pages = [
            slide["page_number"]
            for slide in record.get("_slides", [])
            if slide.get("notes") and parse_target_seconds(slide["notes"]) is None
        ]
        if untimed_pages:
            warnings.append(
                f"Imported speaker notes for {record['role']} have no parseable target_seconds on pages "
                + ", ".join(map(str, untimed_pages))
                + "; those pages contribute 0 to total_seconds."
            )

    route = recommend_route(records, args.goal)
    policy = recommend_policy(records, args.source_policy, args.goal)
    status = "failed" if errors else ("passed_with_warnings" if warnings else "passed")
    manifest: dict[str, Any] = {
        "schema_version": 1,
        "tool": "ppt-gen-source-adapter",
        "generated_at": now(),
        "status": status,
        "project_dir": str(project_dir),
        "goal": args.goal,
        "data_handling": args.data_handling,
        "data_handling_constraints": (
            [
                "Do not upload source material to external OCR services.",
                "Do not send source material to remote image-generation services.",
                "Use offline inspection and local/native composition only.",
            ]
            if args.data_handling == "local-only"
            else []
        ),
        "ocr_policy": "offline",
        "external_ocr_allowed": False,
        "ask_for_ocr_token": False,
        "content_policy": policy,
        "source_policy": policy["policy"],
        "primary_source_id": primary_source_id(records),
        "recommended_route": route,
        "inferred_start": route["entry_phase"],
        "authority_order": authority_order(records),
        "inputs": [public_record(record) for record in records],
        "sources": [public_record(record) for record in records],
        "errors": errors,
        "warnings": list(dict.fromkeys(warnings)),
        "ledgers": {},
    }
    manifest_path = project_dir / "source-manifest.json"
    atomic_write_json(manifest_path, manifest)

    if not args.skip_ledgers and not errors:
        manifest["ledgers"] = initialize_ledgers(
            project_dir, records, policy, args.data_handling, args.force_ledgers
        )
        atomic_write_json(manifest_path, manifest)
    elif args.skip_ledgers:
        manifest["ledgers"] = {"status": "skipped_by_request"}
        atomic_write_json(manifest_path, manifest)
    else:
        manifest["ledgers"] = {"status": "skipped_due_to_errors"}
        atomic_write_json(manifest_path, manifest)

    print(json.dumps(manifest, ensure_ascii=False, indent=2))
    return {"passed": 0, "passed_with_warnings": 2, "failed": 1}[status]


if __name__ == "__main__":
    raise SystemExit(main())
