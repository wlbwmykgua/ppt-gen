#!/usr/bin/env python3
"""Validate report-to-PPT delivery artifacts and their machine-readable ledgers.

The legacy command-line flags remain supported.  New ledger flags make page
order, authoritative copy, speaker-note mapping, and editppt evidence
deterministically checkable instead of treating package counts as final QA.
"""

from __future__ import annotations

import argparse
import copy
import datetime as dt
import hashlib
import io
import json
import os
import posixpath
import re
import subprocess
import sys
import unicodedata
import warnings
import zipfile
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path, PurePosixPath
from typing import Any, Iterable, Sequence
from urllib.parse import unquote
from xml.etree import ElementTree as ET

from artifact_contracts import CLI_DELIVERABLE_ROLES, matching_import
from local_runtime import configure_cache, recognize_image, render_pages
from speech_timing import validate_timing, parse_duration_seconds, expected_entry_duration
from offline_prepare import validate_report as validate_offline_preparation

try:
    from project_state import detect_drift as project_state_detect_drift
    from project_state import manifest_errors as project_state_manifest_errors
except ImportError:  # pragma: no cover - only when this script is relocated incorrectly
    project_state_detect_drift = None
    project_state_manifest_errors = None

try:
    from PIL import Image, ImageOps, UnidentifiedImageError
except ImportError:  # pragma: no cover - exercised only on an incomplete runtime
    Image = None
    ImageOps = None
    UnidentifiedImageError = OSError


IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg"}
EMBEDDED_RASTER_EXTENSIONS = IMAGE_EXTENSIONS | {".gif", ".bmp", ".tif", ".tiff", ".webp"}
SLIDE_PART_RE = re.compile(r"^ppt/slides/slide(\d+)\.xml$")
NUMBER_PREFIX_RE = re.compile(r"^(\d+)")
SLIDE_HEADING_RE = re.compile(
    r"^\s*(?:(?:第\s*0*(\d+)\s*页)|(?:slide\s*0*(\d+)))\s*[|｜:：\-–—]?\s*(.*?)\s*$",
    re.IGNORECASE,
)

PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
OFFICE_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PRESENTATION_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

P_SLD_ID = f"{{{PRESENTATION_NS}}}sldId"
P_SLD_SIZE = f"{{{PRESENTATION_NS}}}sldSz"
P_SHAPE = f"{{{PRESENTATION_NS}}}sp"
P_PICTURE = f"{{{PRESENTATION_NS}}}pic"
P_GRAPHIC_FRAME = f"{{{PRESENTATION_NS}}}graphicFrame"
P_CONNECTOR = f"{{{PRESENTATION_NS}}}cxnSp"
A_PARAGRAPH = f"{{{DRAWING_NS}}}p"
A_TEXT = f"{{{DRAWING_NS}}}t"
W_PARAGRAPH = f"{{{WORD_NS}}}p"
W_TEXT = f"{{{WORD_NS}}}t"
W_STYLE = f"{{{WORD_NS}}}pStyle"
R_EMBED = f"{{{OFFICE_REL_NS}}}embed"

SLIDE_TEXT_FIELDS = (
    "title",
    "subtitle",
    "body",
    "body_copy",
    "copy",
    "required_text",
    "critical_fields",
    "names",
    "dates",
    "numbers",
    "units",
    "source_footer",
    "footer",
    "labels",
    "chart_text",
    "table_text",
    "data",
)
OBSERVED_TEXT_FIELDS = (
    "observed_text",
    "verified_text",
    "ocr_text",
    "extracted_text",
    "observed_title",
)
NOTE_TEXT_FIELDS = (
    "required_text",
    "goal",
    "script",
    "speaker_notes",
    "spoken_script",
    "transition",
    "source_caution",
    "delivery_cue",
    "cue",
)

CLAIM_KINDS = {"fact", "analysis", "user-provided"}
CLAIM_VERIFICATIONS = {"verified", "user-authority", "pending", "unsupported"}
DRAFT_SENTINEL_RE = re.compile(r"\bdraft\s+adapter\s+seed\b", re.IGNORECASE)


@dataclass
class Relationship:
    rel_id: str
    rel_type: str
    target: str
    target_mode: str
    resolved: str | None


@dataclass
class PackageInfo:
    path: Path
    names: set[str]
    roots: dict[str, ET.Element]
    rels: dict[str, dict[str, Relationship]]


@dataclass
class ImageRecord:
    path: Path
    width: int
    height: int
    raw_sha256: str
    pixel_sha256: str
    perceptual_hash: int
    manifest_entry: dict[str, Any] | None = None


@dataclass
class DocxParagraph:
    text: str
    style: str = ""


@dataclass
class DocxInfo:
    path: Path
    paragraphs: list[DocxParagraph]
    text: str


@dataclass
class SlideInfo:
    index: int
    part: str
    text: str
    visible_text: str
    title: str
    native_text_shapes: int
    meaningful_native_objects: int
    picture_count: int
    full_slide_pictures: list[str] = field(default_factory=list)
    picture_fingerprints: dict[str, tuple[str, int]] = field(default_factory=dict)
    semantic_fingerprint: str = ""
    text_objects: dict[str, str] = field(default_factory=dict)
    title_object_id: str = ""


@dataclass
class PptxInfo:
    path: Path
    slides: list[SlideInfo]
    width: int
    height: int
    global_fingerprint: str = ""


def add(result: dict[str, Any], level: str, message: str) -> None:
    if message not in result[level]:
        result[level].append(message)


def local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def normalize_text(value: Any) -> str:
    text = unicodedata.normalize("NFKC", str(value or ""))
    return " ".join(text.replace("\u00a0", " ").split())


def contains_draft_sentinel(value: Any) -> bool:
    if isinstance(value, dict):
        return any(contains_draft_sentinel(item) for item in value.values())
    if isinstance(value, list):
        return any(contains_draft_sentinel(item) for item in value)
    return bool(DRAFT_SENTINEL_RE.search(normalize_text(value)))


def compact_text(value: Any) -> str:
    return re.sub(r"\s+", "", normalize_text(value)).casefold()


def text_contains(actual: str, expected: str) -> bool:
    needle = compact_text(expected)
    return bool(needle) and needle in compact_text(actual)


def text_contains_exact_field(actual: str, expected: str) -> bool:
    haystack = normalize_text(actual).casefold()
    needle = normalize_text(expected).casefold()
    if not needle:
        return False
    start_boundary = r"(?<![\w.])" if needle[0].isascii() and needle[0].isalnum() else ""
    end_boundary = r"(?![\w.])" if needle[-1].isascii() and needle[-1].isalnum() else ""
    return re.search(start_boundary + re.escape(needle) + end_boundary, haystack) is not None


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def source_adapter_directory_hash(path: Path) -> str:
    """Reproduce source_adapter.hash_directory for manifest reconciliation."""
    digest = hashlib.sha256()
    number_re = re.compile(r"(\d+)")

    def natural_key(value: str) -> list[Any]:
        return [int(part) if part.isdigit() else part.casefold() for part in number_re.split(value)]

    entries = list(path.rglob("*"))
    symlinks = [entry for entry in entries if entry.is_symlink()]
    if symlinks:
        relative = ", ".join(sorted(entry.relative_to(path).as_posix() for entry in symlinks))
        raise ValueError(f"source directories may not contain symlinks: {relative}")
    files = sorted(
        (item for item in entries if item.is_file()),
        key=lambda item: natural_key(item.relative_to(path).as_posix()),
    )
    for child in files:
        relative = child.relative_to(path).as_posix()
        size = child.stat().st_size
        digest.update(relative.encode("utf-8"))
        digest.update(b"\0")
        digest.update(str(size).encode("ascii"))
        digest.update(b"\0")
        digest.update(sha256_file(child).encode("ascii"))
        digest.update(b"\n")
    return digest.hexdigest()


def positive_int(value: str) -> int:
    parsed = int(value)
    if parsed <= 0:
        raise argparse.ArgumentTypeError("must be a positive integer")
    return parsed


def natural_sort_key(value: str) -> list[Any]:
    return [int(part) if part.isdigit() else part.casefold() for part in re.split(r"(\d+)", value)]


def unit_interval(value: str) -> float:
    parsed = float(value)
    if not 0.5 <= parsed <= 1.0:
        raise argparse.ArgumentTypeError("must be between 0.5 and 1.0")
    return parsed


def percentage(value: str) -> float:
    parsed = float(value)
    if not 0 <= parsed <= 100:
        raise argparse.ArgumentTypeError("must be between 0 and 100")
    return parsed


def _decoded_image(image: Any) -> tuple[int, int, str, int]:
    if ImageOps is not None:
        image = ImageOps.exif_transpose(image)
    image.load()
    width, height = image.size
    rgba = image.convert("RGBA")
    pixel_digest = hashlib.sha256()
    pixel_digest.update(width.to_bytes(8, "big"))
    pixel_digest.update(height.to_bytes(8, "big"))
    pixel_digest.update(rgba.tobytes())
    reduced = rgba.convert("L").resize((9, 8))
    get_flattened = getattr(reduced, "get_flattened_data", None)
    values = list(get_flattened() if get_flattened is not None else reduced.getdata())
    perceptual = 0
    for row in range(8):
        offset = row * 9
        for col in range(8):
            perceptual = (perceptual << 1) | int(values[offset + col] > values[offset + col + 1])
    return width, height, pixel_digest.hexdigest(), perceptual


def decode_image_path(path: Path) -> tuple[int, int, str, int]:
    if Image is None:
        raise RuntimeError("Pillow is required for full image decoding")
    with warnings.catch_warnings():
        warnings.simplefilter("error", Image.DecompressionBombWarning)
        with Image.open(path) as probe:
            probe.verify()
        with Image.open(path) as decoded:
            return _decoded_image(decoded)


def decode_image_bytes(data: bytes) -> tuple[int, int, str, int]:
    if Image is None:
        raise RuntimeError("Pillow is required for full image decoding")
    with warnings.catch_warnings():
        warnings.simplefilter("error", Image.DecompressionBombWarning)
        with Image.open(io.BytesIO(data)) as probe:
            probe.verify()
        with Image.open(io.BytesIO(data)) as decoded:
            return _decoded_image(decoded)


def hamming_distance(left: int, right: int) -> int:
    return (left ^ right).bit_count()


def image_appears_blank(path: Path) -> bool:
    if Image is None or ImageOps is None:
        return False
    with Image.open(path) as source:
        gray = ImageOps.exif_transpose(source).convert("L").resize((64, 64))
        extrema = gray.getextrema()
        occupied_bins = sum(1 for count in gray.histogram() if count)
        return not extrema or extrema[1] - extrema[0] < 4 or occupied_bins < 4


def offline_ocr_lines(path: Path) -> list[str] | None:
    try:
        lines = recognize_image(path)["lines"]
        return [normalize_text(line) for line in lines if normalize_text(line)]
    except (OSError, subprocess.SubprocessError, ValueError, KeyError, RuntimeError):
        return None


def parse_json_file(path: Path, label: str, result: dict[str, Any]) -> Any | None:
    if not path.is_file():
        add(result, "errors", f"{label} missing or not a file: {path}")
        return None
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as exc:
        add(result, "errors", f"{label} is not valid UTF-8 JSON: {exc}")
        return None


def ledger_entries(data: Any, label: str, result: dict[str, Any]) -> list[dict[str, Any]] | None:
    entries: Any = data
    if isinstance(data, dict):
        if data.get("schema_version") != 1:
            add(result, "errors", f"{label} requires schema_version: 1")
        for key in ("slides", "pages", "entries"):
            if key in data:
                entries = data[key]
                break
        else:
            if data and all(str(key).isdigit() and isinstance(value, dict) for key, value in data.items()):
                entries = [data[key] for key in sorted(data, key=lambda item: int(str(item)))]
            else:
                add(result, "errors", f"{label} must contain a slides/pages/entries array")
                return None
    if not isinstance(entries, list) or not entries:
        add(result, "errors", f"{label} must contain at least one page entry")
        return None
    clean: list[dict[str, Any]] = []
    seen_pages: set[int] = set()
    seen_ids: set[str] = set()
    for index, entry in enumerate(entries, 1):
        if not isinstance(entry, dict):
            add(result, "errors", f"{label} page {index} is not an object")
            continue
        declared = entry_page_number(entry)
        if declared is None:
            add(result, "errors", f"{label} page {index} has no one-based page field")
        elif declared != index:
            add(result, "errors", f"{label} page order mismatch at entry {index}: declares page {declared}")
        elif declared in seen_pages:
            add(result, "errors", f"{label} repeats page {declared}")
        else:
            seen_pages.add(declared)
        entry_id = entry.get("id")
        if isinstance(entry_id, str) and entry_id:
            if entry_id in seen_ids:
                add(result, "errors", f"{label} repeats id {entry_id!r}")
            seen_ids.add(entry_id)
        clean.append(entry)
    return clean


def entry_page_number(entry: dict[str, Any]) -> int | None:
    for key in ("slide", "slide_number", "page", "page_number", "page_index", "index"):
        value = entry.get(key)
        if isinstance(value, int):
            return value
        if isinstance(value, str) and value.strip().isdigit():
            return int(value.strip())
    page_id = entry.get("page_id")
    if isinstance(page_id, str):
        match = re.search(r"(\d+)$", page_id)
        if match:
            return int(match.group(1))
    return None


def flatten_text_values(value: Any) -> list[str]:
    if value is None or isinstance(value, bool):
        return []
    if isinstance(value, (str, int, float)):
        text = normalize_text(value)
        return [text] if text else []
    if isinstance(value, list):
        output: list[str] = []
        for item in value:
            output.extend(flatten_text_values(item))
        return output
    if isinstance(value, dict):
        preferred = ("text", "value", "required_text", "values", "items", "texts")
        if any(key in value for key in preferred):
            output: list[str] = []
            for key in preferred:
                if key in value:
                    output.extend(flatten_text_values(value[key]))
            return output
        output = []
        ignored = {"id", "type", "kind", "decision", "description", "note", "label"}
        for key, item in value.items():
            if key not in ignored:
                output.extend(flatten_text_values(item))
        return output
    return []


def entry_text(entry: dict[str, Any], fields: Sequence[str]) -> list[str]:
    values: list[str] = []
    for field_name in fields:
        if field_name in entry:
            values.extend(flatten_text_values(entry[field_name]))
    unique: list[str] = []
    seen: set[str] = set()
    for value in values:
        key = compact_text(value)
        if key and key not in seen:
            seen.add(key)
            unique.append(value)
    return unique


def load_ledger(path: Path | None, label: str, result: dict[str, Any]) -> tuple[Any, list[dict[str, Any]]] | tuple[None, None]:
    if path is None:
        return None, None
    resolved = path.expanduser().resolve()
    data = parse_json_file(resolved, label, result)
    if data is None:
        return None, None
    entries = ledger_entries(data, label, result)
    if entries is not None:
        validate_ledger_contract(data, entries, label, result)
    return data, entries


def require_entry_key(
    entry: dict[str, Any],
    aliases: Sequence[str],
    label: str,
    page: int,
    result: dict[str, Any],
) -> Any:
    for key in aliases:
        if key in entry:
            return entry[key]
    add(result, "errors", f"{label} page {page} is missing required field {aliases[0]!r}")
    return None


def validate_ledger_contract(
    data: Any,
    entries: list[dict[str, Any]],
    label: str,
    result: dict[str, Any],
) -> None:
    if not isinstance(data, dict):
        return
    if data.get("status") == "draft":
        add(result, "errors", f"{label} is still a draft; complete its authoritative fields before final QA")
    if contains_draft_sentinel(data):
        add(result, "errors", f"{label} still contains adapter draft sentinel text")
    for page, entry in enumerate(entries, 1):
        if label == "slide-copy ledger":
            require_entry_key(entry, ("id",), label, page, result)
            require_entry_key(entry, ("title",), label, page, result)
            require_entry_key(entry, ("body_copy", "body", "copy"), label, page, result)
            critical = require_entry_key(entry, ("critical_fields",), label, page, result)
            if critical is not None and not isinstance(critical, list):
                add(result, "errors", f"{label} page {page} critical_fields must be an array")
            for field_name, aliases in (
                ("purpose", ("purpose",)),
                ("visual_type", ("visual_type", "visual_intent")),
                ("source_ids", ("source_ids",)),
                ("footer", ("footer", "source_footer")),
                ("claim_ids", ("claim_ids",)),
                ("speaker_takeaway", ("speaker_takeaway",)),
            ):
                value = require_entry_key(entry, aliases, label, page, result)
                if field_name in {"source_ids", "claim_ids"} and value is not None and not isinstance(value, list):
                    add(result, "errors", f"{label} page {page} {field_name} must be an array")
            if not normalize_text(entry.get("purpose")):
                add(result, "errors", f"{label} page {page} purpose must express one central judgment")
            if not normalize_text(entry.get("visual_type") or entry.get("visual_intent")):
                add(result, "errors", f"{label} page {page} visual_type must be nonempty")
            if not normalize_text(entry.get("speaker_takeaway")):
                add(result, "errors", f"{label} page {page} speaker_takeaway must be nonempty")
        elif label == "slide manifest":
            require_entry_key(entry, ("id",), label, page, result)
            path_value = require_entry_key(entry, ("path", "file", "image"), label, page, result)
            if isinstance(path_value, str) and not Path(path_value).expanduser().is_absolute():
                add(result, "errors", f"{label} page {page} path must be absolute")
            require_entry_key(entry, ("sha256", "file_sha256"), label, page, result)
            require_entry_key(entry, ("width", "width_px"), label, page, result)
            require_entry_key(entry, ("height", "height_px"), label, page, result)
            require_entry_key(entry, ("title",), label, page, result)
            critical = require_entry_key(entry, ("critical_fields",), label, page, result)
            if critical is not None and not isinstance(critical, list):
                add(result, "errors", f"{label} page {page} critical_fields must be an array")
            observed = require_entry_key(entry, ("observed_text", "verified_text", "ocr_text"), label, page, result)
            if observed is not None and not isinstance(observed, list):
                add(result, "errors", f"{label} page {page} observed_text must be an array")
            verified = require_entry_key(entry, ("text_verified", "copy_ledger_match"), label, page, result)
            if verified is not True:
                add(result, "errors", f"{label} page {page} text_verified must be true for delivery")
            evidence = require_entry_key(entry, ("text_verification_evidence",), label, page, result)
            evidence_path = evidence if isinstance(evidence, str) else evidence.get("path") if isinstance(evidence, dict) else None
            if not isinstance(evidence_path, str) or not Path(evidence_path).expanduser().is_absolute():
                add(result, "errors", f"{label} page {page} text_verification_evidence path must be absolute")
            allow = require_entry_key(
                entry,
                ("allow_full_bleed_image", "allow_image_only", "raster_only_allowed", "full_bleed_photo"),
                label,
                page,
                result,
            )
            if allow is not None and not isinstance(allow, bool):
                add(result, "errors", f"{label} page {page} allow_full_bleed_image must be Boolean")
        elif label == "speaker-notes ledger":
            require_entry_key(entry, ("title",), label, page, result)
            duration = require_entry_key(entry, ("target_seconds", "duration_seconds", "expected_seconds", "duration"), label, page, result)
            parsed_duration = parse_duration_seconds(duration)
            if duration is not None and (parsed_duration is None or parsed_duration <= 0):
                add(result, "errors", f"{label} page {page} target_seconds is invalid")
            for field_name, aliases in (
                ("goal", ("goal",)),
                ("script", ("script", "speaker_notes", "spoken_script")),
                ("cue", ("cue", "delivery_cue")),
                ("transition", ("transition",)),
                ("source_caution", ("source_caution", "fact_or_source_caution")),
                ("claim_ids", ("claim_ids",)),
            ):
                value = require_entry_key(entry, aliases, label, page, result)
                if field_name == "claim_ids" and value is not None and not isinstance(value, list):
                    add(result, "errors", f"{label} page {page} claim_ids must be an array")
                elif field_name != "claim_ids" and not normalize_text(value):
                    add(result, "errors", f"{label} page {page} {field_name} must be nonempty")
    if label == "speaker-notes ledger":
        total = data.get("total_seconds")
        if total is None:
            total = data.get("total_duration_seconds") or data.get("expected_total_seconds")
        parsed_total = parse_duration_seconds(total)
        if parsed_total is None or parsed_total <= 0:
            add(result, "errors", f"{label} is missing a valid total_seconds")


def placeholder_hits(paragraphs: Iterable[str]) -> list[str]:
    hits: list[str] = []
    marker = re.compile(
        r"^\s*(?:[\[【(<{]\s*)?(TODO|TBD|TBC|PLACEHOLDER)(?:\s*[\]】)>}])?\s*(?:[:：\-–—]\s*.*)?$",
        re.IGNORECASE,
    )
    bracketed = re.compile(r"[\[【(<{]\s*(TODO|TBD|TBC|PLACEHOLDER)\s*[\]】)>}]", re.IGNORECASE)
    pending_cn = re.compile(r"(?:^\s*待补充\s*(?:[:：\-–—].*)?$|[\[【（(<{]\s*待补充\s*[\]】）)>}])")
    lorem = re.compile(r"^\s*lorem\s+ipsum(?:\s+dolor\b)?", re.IGNORECASE)
    for paragraph in paragraphs:
        text = unicodedata.normalize("NFKC", paragraph or "").strip()
        if not text:
            continue
        match = marker.match(text) or bracketed.search(text)
        token: str | None = None
        if match:
            token = match.group(1).upper()
        elif pending_cn.search(text):
            token = "待补充"
        elif lorem.search(text):
            token = "Lorem ipsum"
        if token and token not in hits:
            hits.append(token)
    return hits


def rels_source_part(rel_part: str) -> str | None:
    if rel_part == "_rels/.rels":
        return ""
    path = PurePosixPath(rel_part)
    if path.parent.name != "_rels" or not path.name.endswith(".rels"):
        return None
    source_name = path.name[:-5]
    source_parent = path.parent.parent
    return str(source_parent / source_name) if str(source_parent) != "." else source_name


def resolve_relationship_target(source_part: str, target: str) -> str | None:
    clean = unquote(target.split("#", 1)[0].split("?", 1)[0]).replace("\\", "/")
    if not clean:
        return None
    if clean.startswith("/"):
        resolved = posixpath.normpath(clean.lstrip("/"))
    else:
        resolved = posixpath.normpath(posixpath.join(posixpath.dirname(source_part), clean))
    if resolved in {"", "."} or resolved == ".." or resolved.startswith("../"):
        return None
    return resolved


def inspect_opc_package(path: Path, result: dict[str, Any], label: str, main_part: str) -> PackageInfo | None:
    if not path.is_file():
        add(result, "errors", f"{label} missing or not a file: {path}")
        return None
    try:
        with zipfile.ZipFile(path) as archive:
            names_list = archive.namelist()
            names = set(names_list)
            duplicates = sorted(name for name, count in Counter(names_list).items() if count > 1)
            if duplicates:
                add(result, "errors", f"{label} contains duplicate ZIP parts: {duplicates[:8]}")
            unsafe = sorted(
                name for name in names
                if name.startswith("/") or ".." in PurePosixPath(name).parts
            )
            if unsafe:
                add(result, "errors", f"{label} contains invalid package paths: {unsafe[:8]}")
            try:
                bad_crc = archive.testzip()
            except (OSError, RuntimeError, zipfile.BadZipFile) as exc:
                add(result, "errors", f"{label} ZIP CRC/decompression check failed: {exc}")
                return None
            if bad_crc:
                add(result, "errors", f"{label} ZIP CRC check failed for {bad_crc}")

            required = {"[Content_Types].xml", "_rels/.rels", main_part}
            missing_required = sorted(required - names)
            if missing_required:
                add(result, "errors", f"{label} missing required OPC parts: {missing_required}")

            roots: dict[str, ET.Element] = {}
            for name in sorted(names):
                if not (name.endswith(".xml") or name.endswith(".rels")):
                    continue
                try:
                    roots[name] = ET.fromstring(archive.read(name))
                except (ET.ParseError, OSError, RuntimeError, KeyError) as exc:
                    add(result, "errors", f"{label} has malformed XML in {name}: {exc}")

            content_types = roots.get("[Content_Types].xml")
            if content_types is not None:
                for node in content_types.iter():
                    if local_name(node.tag) != "Override":
                        continue
                    part_name = node.attrib.get("PartName", "").lstrip("/")
                    if part_name and part_name not in names:
                        add(result, "errors", f"{label} content types reference missing part: {part_name}")

            rels: dict[str, dict[str, Relationship]] = {}
            for rel_part, root in sorted(roots.items()):
                if not rel_part.endswith(".rels"):
                    continue
                source = rels_source_part(rel_part)
                if source is None:
                    add(result, "errors", f"{label} has invalid relationship part path: {rel_part}")
                    continue
                if source and source not in names:
                    add(result, "errors", f"{label} relationship part has missing source: {rel_part} -> {source}")
                mapping: dict[str, Relationship] = {}
                for node in root.iter():
                    if local_name(node.tag) != "Relationship":
                        continue
                    rel_id = node.attrib.get("Id", "")
                    rel_type = node.attrib.get("Type", "")
                    target = node.attrib.get("Target", "")
                    target_mode = node.attrib.get("TargetMode", "Internal")
                    if not rel_id or not rel_type or not target:
                        add(result, "errors", f"{label} has incomplete relationship in {rel_part}")
                        continue
                    if rel_id in mapping:
                        add(result, "errors", f"{label} has duplicate relationship Id {rel_id} in {rel_part}")
                        continue
                    resolved = None
                    if target_mode.casefold() != "external":
                        resolved = resolve_relationship_target(source, target)
                        if resolved is None or resolved not in names:
                            add(result, "errors", f"{label} relationship target missing: {rel_part}:{rel_id} -> {target}")
                    mapping[rel_id] = Relationship(rel_id, rel_type, target, target_mode, resolved)
                rels[source] = mapping

            root_office_docs = [
                rel for rel in rels.get("", {}).values()
                if rel.rel_type.rstrip("/").endswith("/officeDocument")
            ]
            if not any(rel.resolved == main_part for rel in root_office_docs):
                add(result, "errors", f"{label} root relationships do not target {main_part}")

            for name in sorted(names):
                suffix = Path(name).suffix.lower()
                if suffix in EMBEDDED_RASTER_EXTENSIONS:
                    try:
                        decode_image_bytes(archive.read(name))
                    except Exception as exc:  # Pillow exposes several format-specific exception types.
                        add(result, "errors", f"{label} contains unreadable embedded image {name}: {exc}")
                elif suffix == ".svg":
                    try:
                        ET.fromstring(archive.read(name))
                    except (ET.ParseError, OSError, RuntimeError, KeyError) as exc:
                        add(result, "errors", f"{label} contains malformed SVG {name}: {exc}")

            return PackageInfo(path=path, names=names, roots=roots, rels=rels)
    except (OSError, RuntimeError, zipfile.BadZipFile) as exc:
        add(result, "errors", f"{label} cannot be opened as an OOXML package: {exc}")
        return None


def xml_paragraphs(root: ET.Element, paragraph_tag: str, text_tag: str) -> list[str]:
    paragraphs: list[str] = []
    for paragraph in root.iter(paragraph_tag):
        text = "".join(node.text or "" for node in paragraph.iter(text_tag)).strip()
        if text:
            paragraphs.append(text)
    return paragraphs


def extract_docx_paragraphs(root: ET.Element) -> list[DocxParagraph]:
    output: list[DocxParagraph] = []
    for paragraph in root.iter(W_PARAGRAPH):
        text = "".join(node.text or "" for node in paragraph.iter(W_TEXT)).strip()
        style = ""
        paragraph_properties = paragraph.find(f"{{{WORD_NS}}}pPr")
        if paragraph_properties is not None:
            style_node = paragraph_properties.find(W_STYLE)
            if style_node is not None:
                style = style_node.attrib.get(f"{{{WORD_NS}}}val", "")
        output.append(DocxParagraph(text=text, style=style))
    return output


def validate_docx(path: Path, result: dict[str, Any], label: str) -> DocxInfo | None:
    package = inspect_opc_package(path, result, label, "word/document.xml")
    if package is None:
        return None
    root = package.roots.get("word/document.xml")
    if root is None:
        return None
    paragraphs = extract_docx_paragraphs(root)
    body_text = "\n".join(paragraph.text for paragraph in paragraphs if paragraph.text)
    if not normalize_text(body_text):
        add(result, "errors", f"{label} has no visible body text")

    placeholder_parts = [
        name for name in package.roots
        if name == "word/document.xml"
        or re.fullmatch(r"word/(?:header|footer)\d+\.xml", name)
        or name in {"word/footnotes.xml", "word/endnotes.xml"}
    ]
    for name in sorted(placeholder_parts):
        hits = placeholder_hits(xml_paragraphs(package.roots[name], W_PARAGRAPH, W_TEXT))
        if hits:
            add(result, "errors", f"{label} contains placeholder markers {hits} in {name}")

    if label == "report":
        heading_count = sum(
            1 for paragraph in paragraphs
            if paragraph.text and re.search(r"heading|title|标题", paragraph.style, re.IGNORECASE)
        )
        has_sources = bool(
            re.search(r"(?:https?://|doi\b|参考文献|资料来源|来源[:：]|references?\b|sources?\b)", body_text, re.IGNORECASE)
            or any(rel.rel_type.rstrip("/").endswith("/hyperlink") for rel in package.rels.get("word/document.xml", {}).values())
        )
        result["metrics"]["report_headings"] = heading_count
        result["metrics"]["report_sources_detected"] = has_sources
        if heading_count == 0:
            add(result, "warnings", "report has no paragraph styled as a heading")
        if not has_sources:
            add(result, "warnings", "report has no detectable source section, URL, DOI, or hyperlink")

    result["metrics"][f"{label}_paragraphs"] = sum(1 for paragraph in paragraphs if paragraph.text)
    result["metrics"][f"{label}_characters"] = len(compact_text(body_text))
    return DocxInfo(path=path, paragraphs=paragraphs, text=body_text)


def manifest_file_name(entry: dict[str, Any]) -> str | None:
    for key in ("file", "filename", "image", "path", "source_image", "input"):
        value = entry.get(key)
        if isinstance(value, str) and value.strip():
            return Path(value).name
    return None


def order_slide_images(
    images: list[Path],
    manifest: list[dict[str, Any]] | None,
    result: dict[str, Any],
) -> list[tuple[Path, dict[str, Any] | None]]:
    if manifest is not None:
        by_name: dict[str, Path] = {}
        for image in images:
            if image.name in by_name:
                add(result, "errors", f"duplicate slide image name: {image.name}")
            by_name[image.name] = image
        ordered: list[tuple[Path, dict[str, Any] | None]] = []
        used: set[str] = set()
        for index, entry in enumerate(manifest, 1):
            name = manifest_file_name(entry)
            if not name:
                add(result, "errors", f"slide manifest page {index} has no file/image/path field")
                continue
            if name in used:
                add(result, "errors", f"slide manifest repeats image {name}")
                continue
            image = by_name.get(name)
            if image is None:
                add(result, "errors", f"slide manifest page {index} references missing image {name}")
                continue
            used.add(name)
            ordered.append((image, entry))
        extras = sorted(set(by_name) - used)
        if extras:
            add(result, "errors", f"slide image directory has files absent from manifest: {extras}")
        return ordered

    parsed: list[tuple[int, str, Path]] = []
    missing_prefix: list[str] = []
    for image in images:
        match = NUMBER_PREFIX_RE.match(image.name)
        if match:
            parsed.append((int(match.group(1)), image.name.casefold(), image))
        else:
            missing_prefix.append(image.name)
    if missing_prefix:
        add(
            result,
            "errors",
            f"slide image order is unprovable for unnumbered files {sorted(missing_prefix)}; use numeric prefixes or --slide-manifest",
        )
        return [(image, None) for image in sorted(images, key=lambda item: item.name.casefold())]
    parsed.sort(key=lambda item: (item[0], item[1]))
    numbers = [item[0] for item in parsed]
    duplicates = sorted(number for number, count in Counter(numbers).items() if count > 1)
    if duplicates:
        add(result, "errors", f"slide image numeric prefixes are duplicated: {duplicates}")
    expected = list(range(1, len(parsed) + 1))
    if numbers != expected:
        add(result, "errors", f"slide image numeric sequence must be 1..{len(parsed)}: {numbers}")
    return [(item[2], None) for item in parsed]


def validate_slide_images(
    directory: Path,
    result: dict[str, Any],
    manifest: list[dict[str, Any]] | None,
    min_width: int,
    min_height: int,
) -> list[ImageRecord] | None:
    if not directory.is_dir():
        add(result, "errors", f"slide image directory missing: {directory}")
        return None
    images = [path for path in directory.iterdir() if path.is_file() and path.suffix.lower() in IMAGE_EXTENSIONS]
    if not images:
        add(result, "errors", f"no PNG/JPEG slide images found in {directory}")
        return None
    ordered = order_slide_images(images, manifest, result)
    records: list[ImageRecord] = []
    dimensions: list[tuple[int, int]] = []
    for index, (image, entry) in enumerate(ordered, 1):
        try:
            width, height, pixel_hash, perceptual_hash = decode_image_path(image)
            raw_hash = sha256_file(image)
        except Exception as exc:
            add(result, "errors", f"cannot fully decode image {image.name}: {exc}")
            continue
        dimensions.append((width, height))
        if width * 9 != height * 16:
            add(result, "errors", f"image is not exact 16:9: {image.name} ({width}x{height})")
        if width < min_width or height < min_height:
            add(result, "errors", f"image resolution is below {min_width}x{min_height}: {image.name} ({width}x{height})")
        try:
            if image_appears_blank(image):
                add(result, "errors", f"slide image appears blank or near-uniform: {image.name}")
        except Exception as exc:
            add(result, "errors", f"slide image blank-page inspection failed for {image.name}: {exc}")
        if entry is not None:
            manifest_path_value = entry.get("path") or entry.get("file") or entry.get("image")
            if isinstance(manifest_path_value, str):
                manifest_path = resolve_local_path(directory, manifest_path_value)
                if manifest_path != image.resolve():
                    add(result, "errors", f"slide manifest path mismatch on page {index}: {manifest_path} != {image.resolve()}")
            expected_hash = entry.get("sha256") or entry.get("file_sha256")
            if expected_hash:
                if str(expected_hash).casefold() != raw_hash.casefold():
                    add(result, "errors", f"slide manifest SHA-256 mismatch on page {index}: {image.name}")
            else:
                add(result, "warnings", f"slide manifest page {index} has no SHA-256 binding")
            expected_width = entry.get("width") or entry.get("width_px")
            expected_height = entry.get("height") or entry.get("height_px")
            try:
                if expected_width is not None and int(expected_width) != width:
                    add(result, "errors", f"slide manifest width mismatch on page {index}: {expected_width} != {width}")
                if expected_height is not None and int(expected_height) != height:
                    add(result, "errors", f"slide manifest height mismatch on page {index}: {expected_height} != {height}")
            except (TypeError, ValueError):
                add(result, "errors", f"slide manifest page {index} has nonnumeric dimensions")
            if entry.get("qa_passed") is False or entry.get("passed") is False:
                add(result, "errors", f"slide manifest page {index} records failed QA")
        records.append(ImageRecord(image, width, height, raw_hash, pixel_hash, perceptual_hash, entry))
    if len(set(dimensions)) > 1:
        add(result, "warnings", f"slide images use mixed dimensions: {sorted(set(dimensions))}")
    result["metrics"]["slide_images"] = len(images)
    result["metrics"]["slide_image_dimensions"] = [list(item) for item in sorted(set(dimensions))]
    return records


def validate_manifest_images(
    manifest_path: Path,
    entries: list[dict[str, Any]],
    result: dict[str, Any],
    min_width: int,
    min_height: int,
) -> list[ImageRecord]:
    base = manifest_path.expanduser().resolve().parent
    records: list[ImageRecord] = []
    for index, entry in enumerate(entries, 1):
        path_value = entry.get("path") or entry.get("file") or entry.get("image")
        if not isinstance(path_value, str):
            continue
        image = resolve_local_path(base, path_value)
        if not image.is_file():
            add(result, "errors", f"slide manifest page {index} image is missing: {image}")
            continue
        try:
            width, height, pixel_hash, perceptual_hash = decode_image_path(image)
            raw_hash = sha256_file(image)
        except Exception as exc:
            add(result, "errors", f"slide manifest page {index} image cannot be fully decoded: {exc}")
            continue
        if width * 9 != height * 16:
            add(result, "errors", f"slide manifest image is not exact 16:9 on page {index}: {width}x{height}")
        if width < min_width or height < min_height:
            add(result, "errors", f"slide manifest image is below {min_width}x{min_height} on page {index}")
        try:
            if image_appears_blank(image):
                add(result, "errors", f"slide manifest image appears blank or near-uniform on page {index}")
        except Exception as exc:
            add(result, "errors", f"slide manifest image blank-page inspection failed on page {index}: {exc}")
        expected_hash = entry.get("sha256") or entry.get("file_sha256")
        if not isinstance(expected_hash, str) or expected_hash.casefold() != raw_hash.casefold():
            add(result, "errors", f"slide manifest SHA-256 mismatch on page {index}")
        try:
            expected_width = entry.get("width") if entry.get("width") is not None else entry.get("width_px")
            expected_height = entry.get("height") if entry.get("height") is not None else entry.get("height_px")
            if int(expected_width) != width or int(expected_height) != height:
                add(result, "errors", f"slide manifest dimensions mismatch on page {index}")
        except (TypeError, ValueError):
            add(result, "errors", f"slide manifest page {index} has invalid dimensions")
        records.append(ImageRecord(image, width, height, raw_hash, pixel_hash, perceptual_hash, entry))
    result["metrics"]["slide_manifest_images"] = len(records)
    return records


def relationship_id(node: ET.Element) -> str | None:
    for key, value in node.attrib.items():
        if key.startswith("{") and key.endswith("}id"):
            return value
    return None


def ordered_slide_parts(package: PackageInfo, result: dict[str, Any], label: str) -> list[str]:
    root = package.roots.get("ppt/presentation.xml")
    if root is None:
        return []
    presentation_rels = package.rels.get("ppt/presentation.xml", {})
    parts: list[str] = []
    seen_ids: set[str] = set()
    for node in root.iter(P_SLD_ID):
        rel_id = relationship_id(node)
        if not rel_id:
            add(result, "errors", f"{label} slide id has no relationship id")
            continue
        if rel_id in seen_ids:
            add(result, "errors", f"{label} repeats slide relationship {rel_id}")
            continue
        seen_ids.add(rel_id)
        rel = presentation_rels.get(rel_id)
        if rel is None or not rel.rel_type.rstrip("/").endswith("/slide") or not rel.resolved:
            add(result, "errors", f"{label} slide relationship is missing or not a slide: {rel_id}")
            continue
        if rel.resolved in parts:
            add(result, "errors", f"{label} references slide part twice: {rel.resolved}")
            continue
        parts.append(rel.resolved)
    if not parts:
        add(result, "errors", f"{label} has zero slides")
    discovered = {name for name in package.names if SLIDE_PART_RE.fullmatch(name)}
    orphaned = sorted(discovered - set(parts))
    if orphaned:
        add(result, "errors", f"{label} contains orphan slide parts not in presentation order: {orphaned}")
    missing_xml = [part for part in parts if part not in package.roots]
    if missing_xml:
        add(result, "errors", f"{label} ordered slide XML cannot be parsed: {missing_xml}")
    return parts


def ppt_paragraphs(root: ET.Element) -> list[str]:
    paragraphs = xml_paragraphs(root, A_PARAGRAPH, A_TEXT)
    if paragraphs:
        return paragraphs
    text = "".join(node.text or "" for node in root.iter(A_TEXT)).strip()
    return [text] if text else []


def slide_related_text(package: PackageInfo, slide_part: str) -> list[str]:
    output: list[str] = []
    for rel in package.rels.get(slide_part, {}).values():
        if not rel.resolved:
            continue
        if not (rel.rel_type.rstrip("/").endswith("/chart") or rel.rel_type.rstrip("/").endswith("/diagramData")):
            continue
        root = package.roots.get(rel.resolved)
        if root is None:
            continue
        output.extend(ppt_paragraphs(root))
        output.extend(
            normalize_text(node.text)
            for node in root.iter()
            if local_name(node.tag) == "v" and normalize_text(node.text)
        )
    return output


def slide_title(root: ET.Element, paragraphs: list[str]) -> str:
    for shape in root.iter(P_SHAPE):
        is_title = any(
            local_name(node.tag) == "ph" and node.attrib.get("type") in {"title", "ctrTitle"}
            for node in shape.iter()
        )
        if is_title:
            text = "".join(node.text or "" for node in shape.iter(A_TEXT)).strip()
            if text:
                return text
    return paragraphs[0] if paragraphs else ""


def slide_text_objects(root: ET.Element) -> tuple[dict[str, str], str]:
    """Extract stable OOXML shape identities for revision checkpoint binding."""
    output: dict[str, str] = {}
    title_id = ""
    for shape in root.iter(P_SHAPE):
        non_visual = next(
            (node for node in shape.iter() if local_name(node.tag) == "cNvPr"),
            None,
        )
        text = "\n".join(xml_paragraphs(shape, A_PARAGRAPH, A_TEXT))
        if non_visual is None or not normalize_text(text):
            continue
        shape_id = normalize_text(non_visual.attrib.get("id"))
        shape_name = normalize_text(non_visual.attrib.get("name"))
        canonical = f"shape-{shape_id}" if shape_id else shape_name
        if canonical:
            output[canonical] = text
        if shape_id:
            output[f"id:{shape_id}"] = text
        if shape_name:
            output[f"name:{shape_name}"] = text
        is_title = any(
            local_name(node.tag) == "ph" and node.attrib.get("type") in {"title", "ctrTitle"}
            for node in shape.iter()
        )
        if is_title:
            title_id = canonical
            output["title"] = text
    return output, title_id


def shape_hidden(node: ET.Element) -> bool:
    for child in node.iter():
        if local_name(child.tag) == "cNvPr" and str(child.attrib.get("hidden", "")).casefold() in {"1", "true"}:
            return True
    return False


def node_or_ancestor_hidden(node: ET.Element, parents: dict[ET.Element, ET.Element]) -> bool:
    current: ET.Element | None = node
    while current is not None:
        if shape_hidden(current):
            return True
        current = parents.get(current)
    return False


def shape_bbox(node: ET.Element) -> tuple[int, int, int, int] | None:
    for transform in node.iter():
        if local_name(transform.tag) != "xfrm":
            continue
        offset = next((child for child in transform if local_name(child.tag) == "off"), None)
        extent = next((child for child in transform if local_name(child.tag) == "ext"), None)
        if offset is None or extent is None:
            continue
        try:
            return (
                int(offset.attrib.get("x", "0")),
                int(offset.attrib.get("y", "0")),
                int(extent.attrib.get("cx", "0")),
                int(extent.attrib.get("cy", "0")),
            )
        except ValueError:
            return None
    return None


def bbox_coverage(box: tuple[int, int, int, int] | None, width: int, height: int) -> float:
    if box is None or width <= 0 or height <= 0:
        return 0.0
    x, y, cx, cy = box
    if cx <= 0 or cy <= 0:
        return 0.0
    left, top = max(0, x), max(0, y)
    right, bottom = min(width, x + cx), min(height, y + cy)
    if right <= left or bottom <= top:
        return 0.0
    return ((right - left) * (bottom - top)) / (width * height)


def bbox_is_meaningful_text(box: tuple[int, int, int, int] | None, width: int, height: int) -> bool:
    """Reject one-pixel/off-slide text used only to make a raster look editable."""
    if box is None or width <= 0 or height <= 0:
        return False
    _, _, cx, cy = box
    return (
        cx >= width * 0.005
        and cy >= height * 0.005
        and bbox_coverage(box, width, height) >= 1e-4
    )


def shape_has_visible_geometry(node: ET.Element) -> bool:
    geometry = any(local_name(child.tag) in {"prstGeom", "custGeom"} for child in node.iter())
    if not geometry:
        return False
    no_fills = [child for child in node.iter() if local_name(child.tag) == "noFill"]
    alpha_values = [
        child.attrib.get("val") for child in node.iter()
        if local_name(child.tag) == "alpha" and child.attrib.get("val") is not None
    ]
    if alpha_values and all(str(value) == "0" for value in alpha_values):
        return False
    return len(no_fills) < 2


def picture_part(node: ET.Element, relationships: dict[str, Relationship]) -> str | None:
    for child in node.iter():
        if local_name(child.tag) != "blip":
            continue
        rel_id = child.attrib.get(R_EMBED)
        if rel_id and rel_id in relationships:
            return relationships[rel_id].resolved
    return None


def embedded_fingerprint(package: PackageInfo, part: str) -> tuple[str, int] | None:
    try:
        with zipfile.ZipFile(package.path) as archive:
            _, _, pixel_hash, perceptual_hash = decode_image_bytes(archive.read(part))
        return pixel_hash, perceptual_hash
    except Exception:
        return None


def _relationship_kind(value: str) -> str:
    return value.rstrip("/").rsplit("/", 1)[-1]


def _part_content_fingerprint(
    package: PackageInfo,
    part: str,
    seen: frozenset[str] = frozenset(),
) -> str:
    """Return a content identity that is stable across harmless rId renumbering."""
    if part in seen:
        return f"cycle:{part}"
    suffix = Path(part).suffix.lower()
    if suffix in EMBEDDED_RASTER_EXTENSIONS:
        fingerprint = embedded_fingerprint(package, part)
        if fingerprint is not None:
            return f"pixels:{fingerprint[0]}"
    root = package.roots.get(part)
    if root is not None:
        digest = hashlib.sha256()
        _update_semantic_xml_digest(
            digest,
            root,
            package,
            package.rels.get(part, {}),
            seen | {part},
        )
        return f"xml:{digest.hexdigest()}"
    try:
        with zipfile.ZipFile(package.path) as archive:
            return f"bytes:{sha256_bytes(archive.read(part))}"
    except (OSError, KeyError, zipfile.BadZipFile):
        return f"missing:{part}"


def _update_semantic_xml_digest(
    digest: Any,
    node: ET.Element,
    package: PackageInfo,
    relationships: dict[str, Relationship],
    seen: frozenset[str],
) -> None:
    """Hash semantic XML while ignoring volatile extension metadata."""
    name = local_name(node.tag)
    if name == "extLst":
        return
    digest.update(f"<{name}".encode("utf-8"))
    attributes: list[tuple[str, str]] = []
    for key, raw_value in node.attrib.items():
        attribute_name = local_name(key)
        if attribute_name in {"modId", "creationId"}:
            continue
        relationship = relationships.get(raw_value)
        if relationship is not None:
            if relationship.target_mode.casefold() == "external":
                value = f"external:{_relationship_kind(relationship.rel_type)}:{relationship.target}"
            elif relationship.resolved:
                value = (
                    f"internal:{_relationship_kind(relationship.rel_type)}:"
                    f"{_part_content_fingerprint(package, relationship.resolved, seen)}"
                )
            else:
                value = f"broken:{_relationship_kind(relationship.rel_type)}:{relationship.target}"
        else:
            value = raw_value
        attributes.append((attribute_name, value))
    for attribute_name, value in sorted(attributes):
        digest.update(f"|{attribute_name}={value}".encode("utf-8"))
    digest.update(b">")
    text = normalize_text(node.text)
    if text:
        digest.update(text.encode("utf-8"))
    for child in node:
        _update_semantic_xml_digest(digest, child, package, relationships, seen)
    digest.update(f"</{name}>".encode("utf-8"))


def slide_semantic_fingerprint(package: PackageInfo, root: ET.Element, slide_part: str) -> str:
    digest = hashlib.sha256()
    relationships = package.rels.get(slide_part, {})
    _update_semantic_xml_digest(digest, root, package, relationships, frozenset({slide_part}))
    # Include non-layout/non-notes relationships even if a producer left them
    # unreferenced in the slide XML. This catches silent media/link mutations.
    relationship_records: list[str] = []
    for relationship in relationships.values():
        kind = _relationship_kind(relationship.rel_type)
        if kind in {"slideLayout", "notesSlide"}:
            continue
        if relationship.target_mode.casefold() == "external":
            target_identity = f"external:{relationship.target}"
        elif relationship.resolved:
            target_identity = _part_content_fingerprint(
                package, relationship.resolved, frozenset({slide_part})
            )
        else:
            target_identity = f"broken:{relationship.target}"
        relationship_records.append(f"{kind}:{target_identity}")
    for record in sorted(relationship_records):
        digest.update(b"\0REL\0")
        digest.update(record.encode("utf-8"))
    return digest.hexdigest()


def deck_global_fingerprint(package: PackageInfo, width: int, height: int, slide_count: int) -> str:
    """Hash masters/layouts/themes and global geometry outside per-slide edits."""
    digest = hashlib.sha256()
    digest.update(f"geometry:{width}x{height};slides:{slide_count}".encode("ascii"))
    prefixes = (
        "ppt/slideMasters/",
        "ppt/slideLayouts/",
        "ppt/theme/",
        "ppt/notesMasters/",
    )
    exact = {"ppt/tableStyles.xml", "ppt/presProps.xml"}
    parts = sorted(
        part for part in package.names
        if part in exact or (part.endswith(".xml") and part.startswith(prefixes))
    )
    for part in parts:
        digest.update(b"\0GLOBAL\0")
        digest.update(part.encode("utf-8"))
        digest.update(b"\0")
        digest.update(_part_content_fingerprint(package, part).encode("ascii"))
    return digest.hexdigest()


def analyze_slide_objects(
    package: PackageInfo,
    root: ET.Element,
    slide_part: str,
    width: int,
    height: int,
    full_slide_threshold: float,
) -> tuple[int, int, int, list[str], dict[str, tuple[str, int]], list[str]]:
    relationships = package.rels.get(slide_part, {})
    parents = {child: parent for parent in root.iter() for child in parent}
    native_text_shapes = 0
    meaningful_objects = 0
    picture_count = 0
    full_slide_parts: list[str] = []
    fingerprints: dict[str, tuple[str, int]] = {}
    visible_text: list[str] = []

    for shape in root.iter(P_SHAPE):
        if node_or_ancestor_hidden(shape, parents):
            continue
        text_paragraphs = xml_paragraphs(shape, A_PARAGRAPH, A_TEXT)
        text = "\n".join(text_paragraphs)
        box = shape_bbox(shape)
        inherited_placeholder = box is None and bool(text)
        if text and (bbox_is_meaningful_text(box, width, height) or inherited_placeholder):
            native_text_shapes += 1
            meaningful_objects += 1
            visible_text.extend(text_paragraphs)
        elif bbox_coverage(box, width, height) >= 1e-4 and shape_has_visible_geometry(shape):
            meaningful_objects += 1

    for frame in root.iter(P_GRAPHIC_FRAME):
        if node_or_ancestor_hidden(frame, parents):
            continue
        coverage = bbox_coverage(shape_bbox(frame), width, height)
        has_native_payload = any(local_name(node.tag) in {"tbl", "chart", "relIds"} for node in frame.iter())
        if coverage >= 1e-6 and has_native_payload:
            meaningful_objects += 1
            visible_text.extend(ppt_paragraphs(frame))
            referenced_ids = {
                value
                for node in frame.iter()
                for key, value in node.attrib.items()
                if local_name(key) in {"id", "embed", "link"} and value in relationships
            }
            for rel_id in referenced_ids:
                rel = relationships[rel_id]
                if not rel.resolved or not (
                    rel.rel_type.rstrip("/").endswith("/chart")
                    or rel.rel_type.rstrip("/").endswith("/diagramData")
                ):
                    continue
                related_root = package.roots.get(rel.resolved)
                if related_root is not None:
                    visible_text.extend(ppt_paragraphs(related_root))
                    visible_text.extend(
                        normalize_text(node.text)
                        for node in related_root.iter()
                        if local_name(node.tag) == "v" and normalize_text(node.text)
                    )

    for connector in root.iter(P_CONNECTOR):
        if not node_or_ancestor_hidden(connector, parents) and bbox_coverage(shape_bbox(connector), width, height) >= 1e-8:
            meaningful_objects += 1

    for picture in root.iter(P_PICTURE):
        if node_or_ancestor_hidden(picture, parents):
            continue
        picture_count += 1
        part = picture_part(picture, relationships)
        coverage = bbox_coverage(shape_bbox(picture), width, height)
        if part:
            if part not in fingerprints:
                fingerprint = embedded_fingerprint(package, part)
                if fingerprint is not None:
                    fingerprints[part] = fingerprint
            if coverage >= full_slide_threshold:
                full_slide_parts.append(part)

    for background in root.iter():
        if local_name(background.tag) != "bg":
            continue
        part = picture_part(background, relationships)
        if part:
            picture_count += 1
            if part not in fingerprints:
                fingerprint = embedded_fingerprint(package, part)
                if fingerprint is not None:
                    fingerprints[part] = fingerprint
            full_slide_parts.append(part)

    return native_text_shapes, meaningful_objects, picture_count, full_slide_parts, fingerprints, visible_text


def matches_source_image(fingerprint: tuple[str, int], source: ImageRecord) -> bool:
    pixel_hash, perceptual_hash = fingerprint
    return pixel_hash == source.pixel_sha256 or hamming_distance(perceptual_hash, source.perceptual_hash) <= 4


def compare_slide_copy(
    slides: list[SlideInfo],
    ledger: list[dict[str, Any]],
    result: dict[str, Any],
    label: str,
) -> None:
    if len(slides) != len(ledger):
        add(result, "errors", f"{label} cannot reconcile copy: {len(slides)} slides != {len(ledger)} ledger pages")
        return
    for slide, entry in zip(slides, ledger):
        critical = entry_text(
            entry,
            ("critical_fields", "names", "dates", "numbers", "units"),
        )
        regular = entry_text(
            entry,
            tuple(field for field in SLIDE_TEXT_FIELDS if field not in {"title", "critical_fields", "names", "dates", "numbers", "units"}),
        )
        missing = [value for value in regular if not text_contains(slide.visible_text, value)]
        missing.extend(value for value in critical if not text_contains_exact_field(slide.visible_text, value))
        if missing:
            add(result, "errors", f"{label} slide {slide.index} is missing authoritative text: {missing}")
        expected_title = normalize_text(entry.get("title", ""))
        native_paragraphs = [compact_text(line) for line in slide.visible_text.splitlines() if normalize_text(line)]
        if expected_title and compact_text(expected_title) not in native_paragraphs:
            add(result, "errors", f"{label} slide {slide.index} title mismatch: {slide.title!r} != {expected_title!r}")


def pptx_info(
    path: Path,
    result: dict[str, Any],
    label: str,
    editable: bool,
    ledger: list[dict[str, Any]] | None,
    page_manifest: list[dict[str, Any]] | None,
    source_images: list[ImageRecord] | None,
    full_slide_threshold: float,
) -> PptxInfo | None:
    package = inspect_opc_package(path, result, label, "ppt/presentation.xml")
    if package is None:
        return None
    root = package.roots.get("ppt/presentation.xml")
    if root is None:
        return None
    slide_size = next(root.iter(P_SLD_SIZE), None)
    if slide_size is None:
        add(result, "errors", f"{label} has no slide size")
        width = height = 0
    else:
        try:
            width, height = int(slide_size.attrib["cx"]), int(slide_size.attrib["cy"])
        except (KeyError, ValueError):
            add(result, "errors", f"{label} has invalid slide size attributes")
            width = height = 0
        if width > 0 and height > 0 and width * 9 != height * 16:
            add(result, "errors", f"{label} slide size is not exact 16:9: {width}x{height}")

    parts = ordered_slide_parts(package, result, label)
    if source_images is not None and len(source_images) != len(parts):
        add(
            result,
            "errors",
            f"{label} approved source image count {len(source_images)} != slide count {len(parts)}",
        )
    if page_manifest is not None and len(page_manifest) != len(parts):
        add(
            result,
            "errors",
            f"{label} slide-manifest page count {len(page_manifest)} != slide count {len(parts)}",
        )
    slides: list[SlideInfo] = []
    inventories: list[dict[str, Any]] = []
    for index, part in enumerate(parts, 1):
        slide_root = package.roots.get(part)
        if slide_root is None:
            continue
        paragraphs = ppt_paragraphs(slide_root)
        related = slide_related_text(package, part)
        hits = placeholder_hits(paragraphs + related)
        if hits:
            add(result, "errors", f"{label} contains placeholder markers {hits} on slide {index}")
        text = "\n".join(paragraphs + related)
        title = slide_title(slide_root, paragraphs)
        native_text_shapes, meaningful, picture_count, full_parts, fingerprints, visible_paragraphs = analyze_slide_objects(
            package, slide_root, part, width, height, full_slide_threshold
        )
        semantic_fingerprint = slide_semantic_fingerprint(package, slide_root, part)
        text_objects, title_object_id = slide_text_objects(slide_root)
        visible_text = "\n".join(value for value in visible_paragraphs if normalize_text(value))
        visible_title = slide_title(slide_root, visible_paragraphs)
        slide = SlideInfo(
            index=index,
            part=part,
            text=text,
            visible_text=visible_text,
            title=visible_title or title,
            native_text_shapes=native_text_shapes,
            meaningful_native_objects=meaningful,
            picture_count=picture_count,
            full_slide_pictures=full_parts,
            picture_fingerprints=fingerprints,
            semantic_fingerprint=semantic_fingerprint,
            text_objects=text_objects,
            title_object_id=title_object_id,
        )
        slides.append(slide)
        inventories.append(
            {
                "slide": index,
                "native_text_shapes": native_text_shapes,
                "meaningful_native_objects": meaningful,
                "pictures": picture_count,
                "full_slide_pictures": len(full_parts),
            }
        )

        source = source_images[index - 1] if source_images and index <= len(source_images) else None
        matching_full_parts = [
            part_name
            for part_name in full_parts
            if source
            and part_name in fingerprints
            and matches_source_image(fingerprints[part_name], source)
        ]
        source_match = bool(matching_full_parts)
        image_authority_required = label == "image_pptx" or (
            label in {"reference_pptx", "ppt_notes_pptx"}
            and source is not None
            and bool(full_parts)
        )
        if image_authority_required and source is not None and not source_match:
            add(
                result,
                "errors",
                f"{label} slide {index} does not contain the matching approved full-slide image",
            )
        elif image_authority_required and source is not None and len(full_parts) != 1:
            add(
                result,
                "errors",
                f"{label} slide {index} must contain exactly one approved full-slide image; found {len(full_parts)}",
            )

        if editable:
            entry = page_manifest[index - 1] if page_manifest and index <= len(page_manifest) else {}
            allow_image_only = bool(
                entry.get("allow_full_bleed_image")
                or entry.get("allow_image_only")
                or entry.get("raster_only_allowed")
                or entry.get("full_bleed_photo")
            )
            if source_match:
                add(result, "errors", f"{label} slide {index} embeds the approved source slide as a full-page screenshot")
            if picture_count and meaningful == 0:
                add(result, "errors", f"{label} slide {index} has pictures but no visible meaningful editable objects")
            elif not picture_count and meaningful == 0:
                add(result, "errors", f"{label} slide {index} has no visible meaningful editable objects")
            elif full_parts and not source_match:
                if allow_image_only and meaningful > 0 and ledger is not None:
                    add(result, "warnings", f"{label} slide {index} uses a ledger-approved full-bleed raster with native editable content")
                else:
                    add(result, "errors", f"{label} slide {index} has a full-slide raster without an explicit ledger exception and native content")

    if ledger is not None and editable:
        compare_slide_copy(slides, ledger, result, label)
    result["metrics"][f"{label}_size"] = [width, height]
    result["metrics"][f"{label}_slides"] = len(parts)
    global_fingerprint = deck_global_fingerprint(package, width, height, len(parts))
    result["metrics"][f"{label}_semantic_fingerprints"] = [
        slide.semantic_fingerprint for slide in slides
    ]
    result["metrics"][f"{label}_global_fingerprint"] = global_fingerprint
    if editable:
        result["metrics"][f"{label}_object_inventory"] = inventories
    return PptxInfo(
        path=path,
        slides=slides,
        width=width,
        height=height,
        global_fingerprint=global_fingerprint,
    )


def compare_image_manifest_copy(
    images: list[ImageRecord] | None,
    slide_manifest: list[dict[str, Any]] | None,
    slide_copy: list[dict[str, Any]] | None,
    manifest_base: Path | None,
    result: dict[str, Any],
) -> None:
    if images is None:
        return
    if slide_manifest is None:
        if slide_copy is not None:
            add(result, "errors", "slide image text cannot be reconciled to --slide-copy-ledger without --slide-manifest evidence")
        return
    if len(images) != len(slide_manifest) or (slide_copy is not None and len(slide_manifest) != len(slide_copy)):
        return
    for index, (image, manifest_entry) in enumerate(zip(images, slide_manifest), 1):
        copy_entry = slide_copy[index - 1] if slide_copy is not None else {}
        manifest_title = normalize_text(manifest_entry.get("title", ""))
        expected_title = normalize_text(copy_entry.get("title", "")) if slide_copy is not None else manifest_title
        if expected_title and compact_text(expected_title) != compact_text(manifest_title):
            add(result, "errors", f"slide manifest page {index} title mismatch: {manifest_title!r} != {expected_title!r}")
        manifest_critical = entry_text(manifest_entry, ("critical_fields",))
        expected_critical = entry_text(copy_entry, ("critical_fields",)) if slide_copy is not None else manifest_critical
        missing_critical = [
            value for value in expected_critical
            if not any(compact_text(value) == compact_text(observed) for observed in manifest_critical)
        ]
        extra_critical = [
            value for value in manifest_critical
            if not any(compact_text(value) == compact_text(expected) for expected in expected_critical)
        ]
        if missing_critical or extra_critical:
            add(
                result,
                "errors",
                f"slide manifest page {index} critical_fields mismatch: missing={missing_critical}, extra={extra_critical}",
            )
        required_body = entry_text(copy_entry, ("body_copy", "body", "copy", "footer", "source_footer"))
        observed = "\n".join(entry_text(manifest_entry, OBSERVED_TEXT_FIELDS))
        attested = manifest_entry.get("text_verified") is True or manifest_entry.get("copy_ledger_match") is True
        if attested:
            expected_hash = manifest_entry.get("sha256") or manifest_entry.get("file_sha256")
            if not expected_hash:
                add(result, "errors", f"slide manifest page {index} attests text without binding the image SHA-256")
            elif str(expected_hash).casefold() != image.raw_sha256.casefold():
                add(result, "errors", f"slide manifest page {index} text attestation is bound to the wrong image")
            evidence = manifest_entry.get("text_verification_evidence")
            evidence_path_value: str | None = None
            evidence_image_hash: str | None = None
            if isinstance(evidence, str):
                evidence_path_value = evidence
            elif isinstance(evidence, dict):
                path_value = evidence.get("path") or evidence.get("file") or evidence.get("record")
                if isinstance(path_value, str):
                    evidence_path_value = path_value
                hash_value = evidence.get("image_sha256") or evidence.get("source_image_sha256") or evidence.get("slide_sha256")
                if hash_value is not None:
                    evidence_image_hash = str(hash_value)
            if not evidence_path_value or manifest_base is None:
                add(result, "errors", f"slide manifest page {index} text attestation has no text_verification_evidence path")
            else:
                evidence_path = resolve_local_path(manifest_base, evidence_path_value)
                if not evidence_path.is_file():
                    add(result, "errors", f"slide manifest page {index} text verification evidence is missing: {evidence_path}")
                else:
                    if evidence_path.suffix.lower() != ".json":
                        add(result, "errors", f"slide page {index} text evidence must be an independent JSON review record")
                    else:
                        evidence_data = parse_json_file(evidence_path, f"slide page {index} text evidence", result)
                        if isinstance(evidence_data, dict):
                            if evidence_data.get("schema_version") != 1:
                                add(result, "errors", f"slide page {index} text evidence requires schema_version=1")
                            if evidence_data.get("tool") not in {"ppt-gen.slide-text-review", "offline-ocr-review"}:
                                add(result, "errors", f"slide page {index} text evidence tool is not approved")
                            if evidence_data.get("passed") is not True:
                                add(result, "errors", f"slide page {index} text evidence did not pass")
                            reviewer = evidence_data.get("reviewer") if isinstance(evidence_data.get("reviewer"), dict) else {}
                            if not normalize_text(reviewer.get("method")):
                                add(result, "errors", f"slide page {index} text evidence lacks reviewer.method")
                            if reviewer.get("ocr_policy") not in {None, "offline"}:
                                add(result, "errors", f"slide page {index} text evidence violates offline OCR policy")
                            source_image_path = evidence_data.get("image_path") or evidence_data.get("source_image_path")
                            if not isinstance(source_image_path, str) or resolve_local_path(evidence_path.parent, source_image_path) != image.path.resolve():
                                add(result, "errors", f"slide page {index} text evidence image_path does not match the slide image")
                            evidence_image_hash = str(
                                evidence_data.get("image_sha256")
                                or evidence_data.get("source_image_sha256")
                                or evidence_data.get("slide_sha256")
                                or evidence_image_hash
                                or ""
                            )
                            evidence_observed = entry_text(evidence_data, OBSERVED_TEXT_FIELDS)
                            if not evidence_observed:
                                add(result, "errors", f"slide page {index} text evidence requires observed_text")
                            else:
                                observed = "\n".join(evidence_observed)
                    if not evidence_image_hash:
                        add(result, "errors", f"slide page {index} text evidence is not bound to the current image SHA-256")
                    elif evidence_image_hash.casefold() != image.raw_sha256.casefold():
                        add(result, "errors", f"slide page {index} text evidence image SHA-256 mismatch")
        if observed:
            missing = [value for value in required_body if not text_contains(observed, value)]
            missing.extend(value for value in expected_critical if not text_contains_exact_field(observed, value))
            if expected_title and not text_contains(observed, expected_title):
                missing.insert(0, expected_title)
            if missing:
                add(result, "errors", f"slide image page {index} is missing authoritative text in manifest evidence: {missing}")
        elif required_body:
            add(
                result,
                "warnings",
                f"slide manifest page {index} binds title/critical fields but has no observed_text or copy_ledger_match attestation for body copy",
            )
        # The JSON review is provenance, not self-authenticating content.  On
        # macOS, independently re-read the actual pixels with Apple Vision's
        # fully local OCR and reconcile every authoritative string.  If that
        # local verifier is unavailable, final image delivery fails closed.
        independent_text = offline_ocr_lines(image.path)
        if independent_text is None:
            add(result, "errors", f"slide image page {index} could not run independent offline OCR")
        else:
            independent_observed = "\n".join(independent_text)
            missing_independent = [
                value for value in required_body
                if not text_contains(independent_observed, value)
            ]
            missing_independent.extend(
                value for value in expected_critical
                if not text_contains_exact_field(independent_observed, value)
            )
            if expected_title and not text_contains(independent_observed, expected_title):
                missing_independent.insert(0, expected_title)
            if missing_independent:
                add(
                    result,
                    "errors",
                    f"slide image page {index} independent offline OCR is missing authoritative text: {missing_independent}",
                )


def duration_from_paragraphs(paragraphs: Sequence[str], require_label: bool = True) -> int | None:
    for paragraph in paragraphs:
        if require_label and not re.search(r"(?:时长|time|duration)", paragraph, re.IGNORECASE):
            continue
        parsed = parse_duration_seconds(paragraph)
        if parsed is not None:
            return parsed
    return None


def validate_speaker_notes(
    info: DocxInfo | None,
    result: dict[str, Any],
    expected_count: int | None,
    notes_data: Any,
    notes_ledger: list[dict[str, Any]] | None,
    title_entries: list[dict[str, Any]] | None,
    deck_titles: list[str] | None,
    tolerance_percent: float,
) -> int | None:
    if info is None:
        return None
    sections: list[dict[str, Any]] = []
    current: dict[str, Any] | None = None
    cover_paragraphs: list[str] = []
    for paragraph in info.paragraphs:
        if not paragraph.text:
            continue
        match = SLIDE_HEADING_RE.match(paragraph.text)
        if match:
            page = int(match.group(1) or match.group(2))
            current = {"page": page, "title": normalize_text(match.group(3)), "paragraphs": []}
            sections.append(current)
        elif current is None:
            cover_paragraphs.append(paragraph.text)
        else:
            current["paragraphs"].append(paragraph.text)
    if not sections:
        add(result, "errors", "speaker_notes has no recognizable '第 N 页' or 'Slide N' sections")
        return 0
    page_numbers = [section["page"] for section in sections]
    expected_numbers = list(range(1, len(sections) + 1))
    if page_numbers != expected_numbers:
        add(result, "errors", f"speaker_notes page sequence must be 1..{len(sections)}: {page_numbers}")
    if expected_count is not None and len(sections) != expected_count:
        add(result, "errors", f"speaker_notes sections {len(sections)} != expected slides {expected_count}")

    expected_titles: list[str] = []
    strict_title_match = False
    if title_entries is not None:
        expected_titles = [normalize_text(entry.get("title", "")) for entry in title_entries]
        strict_title_match = True
    elif deck_titles is not None:
        expected_titles = [normalize_text(title) for title in deck_titles]
    if expected_titles and len(expected_titles) == len(sections):
        for section, expected_title in zip(sections, expected_titles):
            actual_key = compact_text(section["title"])
            expected_key = compact_text(expected_title)
            matches = (
                actual_key == expected_key
                if strict_title_match
                else expected_key in actual_key or actual_key in expected_key
            )
            if expected_title and not matches:
                add(
                    result,
                    "errors",
                    f"speaker_notes page {section['page']} title mismatch: {section['title']!r} != {expected_title!r}",
                )
    elif expected_titles:
        add(result, "errors", f"speaker_notes title mapping cannot be reconciled: {len(sections)} sections != {len(expected_titles)} expected titles")
    else:
        add(result, "warnings", "speaker_notes titles could not be reconciled because no slide deck or title ledger was provided")

    durations: list[int] = []
    for index, section in enumerate(sections, 1):
        duration = duration_from_paragraphs(section["paragraphs"][:5])
        if duration is None:
            add(result, "errors", f"speaker_notes page {index} has no labeled duration")
            durations.append(0)
        elif duration <= 0:
            add(result, "errors", f"speaker_notes page {index} duration must be positive")
            durations.append(duration)
        else:
            durations.append(duration)

    if notes_ledger is not None:
        if len(notes_ledger) != len(sections):
            add(result, "errors", f"speaker-notes ledger pages {len(notes_ledger)} != speaker_notes sections {len(sections)}")
        else:
            for index, (section, ledger_entry, actual_duration) in enumerate(zip(sections, notes_ledger, durations), 1):
                expected_title = normalize_text(ledger_entry.get("title", ""))
                if expected_title and compact_text(section["title"]) != compact_text(expected_title):
                    add(result, "errors", f"speaker-notes ledger title mismatch on page {index}")
                section_text = "\n".join(section["paragraphs"])
                missing = [value for value in entry_text(ledger_entry, NOTE_TEXT_FIELDS) if not text_contains(section_text, value)]
                if missing:
                    add(result, "errors", f"speaker_notes page {index} is missing ledger text: {missing}")
                expected_duration = expected_entry_duration(ledger_entry)
                if expected_duration is not None and expected_duration > 0:
                    drift = abs(actual_duration - expected_duration) / expected_duration * 100
                    if drift > tolerance_percent:
                        add(result, "errors", f"speaker_notes page {index} duration differs by {drift:.1f}%")

    actual_total = sum(durations)
    stated_total = duration_from_paragraphs(
        [paragraph for paragraph in cover_paragraphs if re.search(r"总时长|total\s+(?:time|duration)", paragraph, re.IGNORECASE)]
    )
    ledger_total = None
    if isinstance(notes_data, dict):
        for key in ("total_seconds", "total_duration_seconds", "expected_total_seconds", "total_duration", "total_time"):
            if key in notes_data:
                ledger_total = parse_duration_seconds(notes_data[key])
                break
    target_total = ledger_total if ledger_total is not None else stated_total
    if target_total and target_total > 0:
        drift = abs(actual_total - target_total) / target_total * 100
        if drift > tolerance_percent:
            add(result, "errors", f"speaker_notes total duration {actual_total}s differs from stated/ledger total {target_total}s by {drift:.1f}%")
    result["metrics"]["speaker_notes_sections"] = len(sections)
    result["metrics"]["speaker_notes_total_seconds"] = actual_total
    return len(sections)


def validate_speaker_notes_authorities(
    notes_data: Any,
    notes_ledger_path: Path | None,
    slide_copy_path: Path | None,
    reference_pptx_path: Path | None,
    result: dict[str, Any],
) -> None:
    if notes_ledger_path is None:
        return
    if not isinstance(notes_data, dict):
        return
    authorities = notes_data.get("reference_authorities")
    if not isinstance(authorities, dict):
        add(result, "errors", "speaker-notes ledger requires reference_authorities")
        return
    for key, expected in (
        ("slide_copy_ledger", slide_copy_path),
        ("reference_deck", reference_pptx_path),
    ):
        record = authorities.get(key)
        if not isinstance(record, dict):
            add(result, "errors", f"speaker-notes reference_authorities.{key} must be an object")
            continue
        path_value = record.get("path")
        if expected is None:
            add(result, "errors", f"speaker-notes {key} authority cannot be reconciled without its command-line artifact")
            continue
        expected_path = expected.expanduser().resolve()
        if not isinstance(path_value, str) or not Path(path_value).expanduser().is_absolute():
            add(result, "errors", f"speaker-notes reference_authorities.{key}.path must be absolute")
        elif Path(path_value).expanduser().resolve() != expected_path:
            add(result, "errors", f"speaker-notes reference_authorities.{key}.path does not match the current artifact")
        if not expected_path.is_file():
            add(result, "errors", f"speaker-notes {key} authority is missing: {expected_path}")
            continue
        expected_hash = record.get("sha256")
        if not isinstance(expected_hash, str) or expected_hash.casefold() != sha256_file(expected_path).casefold():
            add(result, "errors", f"speaker-notes reference_authorities.{key}.sha256 is missing or mismatched")


def resolve_local_path(base: Path, value: str) -> Path:
    candidate = Path(value).expanduser()
    return candidate.resolve() if candidate.is_absolute() else (base / candidate).resolve()


def first_existing(paths: Iterable[Path]) -> Path | None:
    for path in paths:
        if path.is_file():
            return path
    return None


def load_bound_validation_record(
    value: Any,
    base: Path,
    label: str,
    result: dict[str, Any],
    fallback_hash: Any = None,
) -> dict[str, Any] | None:
    record_path_value: str | None = None
    expected_hash = fallback_hash
    if isinstance(value, str):
        record_path_value = value
    elif isinstance(value, dict):
        candidate = value.get("path") or value.get("file") or value.get("validation")
        if isinstance(candidate, str):
            record_path_value = candidate
        expected_hash = value.get("sha256") or expected_hash
    if not record_path_value:
        add(result, "errors", f"{label} must reference a validation JSON file")
        return None
    record_path = resolve_local_path(base, record_path_value)
    if not record_path.is_file():
        add(result, "errors", f"{label} validation file is missing: {record_path}")
        return None
    actual_hash = sha256_file(record_path)
    if not isinstance(expected_hash, str):
        add(result, "errors", f"{label} validation file has no SHA-256 binding")
    elif expected_hash.casefold() != actual_hash.casefold():
        add(result, "errors", f"{label} validation file SHA-256 mismatch")
    record = parse_json_file(record_path, label, result)
    if not isinstance(record, dict) or record.get("passed") is not True:
        add(result, "errors", f"{label} validation did not pass")
        return None
    if record.get("status") == "failed" or (isinstance(record.get("errors"), list) and record["errors"]):
        add(result, "errors", f"{label} validation is internally inconsistent with recorded errors")
    return record


def validate_bound_evidence_file(
    value: Any,
    base: Path,
    label: str,
    result: dict[str, Any],
    *,
    minimum_image_width: int = 1280,
    minimum_image_height: int = 720,
    allowed_suffixes: set[str] | None = None,
) -> Path | None:
    path_value = value
    expected_hash = None
    if isinstance(value, dict):
        path_value = value.get("path") or value.get("file")
        expected_hash = value.get("sha256")
    if not isinstance(path_value, str):
        add(result, "errors", f"{label} must include an evidence path")
        return None
    path = resolve_local_path(base, path_value)
    if not path.is_file() or path.stat().st_size == 0:
        add(result, "errors", f"{label} is missing or empty: {path}")
        return None
    if allowed_suffixes is not None and path.suffix.lower() not in allowed_suffixes:
        add(result, "errors", f"{label} must use one of {sorted(allowed_suffixes)}")
    actual_hash = sha256_file(path)
    if not isinstance(expected_hash, str):
        add(result, "errors", f"{label} requires a SHA-256 binding")
    elif expected_hash.casefold() != actual_hash.casefold():
        add(result, "errors", f"{label} SHA-256 mismatch")
    if path.suffix.lower() in IMAGE_EXTENSIONS:
        try:
            width, height, _, _ = decode_image_path(path)
            if width < minimum_image_width or height < minimum_image_height:
                add(result, "errors", f"{label} is below readable size: {width}x{height}")
        except Exception as exc:
            add(result, "errors", f"{label} cannot be decoded: {exc}")
    elif path.suffix.lower() == ".json":
        payload = parse_json_file(path, label, result)
        if not isinstance(payload, dict) or payload.get("passed") is not True:
            add(result, "errors", f"{label} JSON evidence did not pass")
        elif payload.get("status") == "failed" or (isinstance(payload.get("errors"), list) and payload["errors"]):
            add(result, "errors", f"{label} JSON evidence is internally failed")
    else:
        add(result, "errors", f"{label} must be readable PNG/JPEG or passed JSON evidence, not {path.suffix or '(no extension)'}")
    return path


def validate_render_evidence_path(
    path: Path,
    artifact_path: Path | None,
    label: str,
    result: dict[str, Any],
    expected_pages: int | None = None,
) -> None:
    resolved = path.expanduser().resolve()
    if not resolved.is_file() or resolved.suffix.lower() != ".json":
        add(result, "errors", f"{label} must be a hash-bound render manifest JSON: {resolved}")
        return
    data = parse_json_file(resolved, label, result)
    if not isinstance(data, dict):
        return
    if data.get("schema_version") != 1 or data.get("passed") is not True:
        add(result, "errors", f"{label} requires schema_version=1 and passed=true")
    if data.get("tool") != "ppt-gen.render-evidence":
        add(result, "errors", f"{label} tool must be ppt-gen.render-evidence")
    renderer = data.get("renderer") if isinstance(data.get("renderer"), dict) else {}
    renderer_name = renderer.get("name")
    trusted_renderers = {
        "documents:documents",
        "pptx",
        "presentations:Presentations",
        "libreoffice",
        "microsoft-powerpoint",
        "keynote",
        "wps-office",
        "image-to-editable-ppt",
    }
    if renderer_name not in trusted_renderers:
        add(result, "errors", f"{label} renderer.name must identify an approved document/PPT renderer")
    if not normalize_text(renderer.get("version")) and not normalize_text(renderer.get("implementation")):
        add(result, "errors", f"{label} renderer requires version or implementation provenance")
    if artifact_path is None:
        add(result, "errors", f"{label} has no artifact to bind")
        return
    artifact = artifact_path.expanduser().resolve()
    artifact_value = data.get("artifact") if isinstance(data.get("artifact"), dict) else {}
    artifact_record_path = artifact_value.get("path")
    artifact_hash = artifact_value.get("sha256")
    if not isinstance(artifact_record_path, str) or Path(artifact_record_path).expanduser().resolve() != artifact:
        add(result, "errors", f"{label} artifact.path does not match the delivered artifact")
    if not artifact.is_file() or not isinstance(artifact_hash, str) or artifact_hash.casefold() != sha256_file(artifact).casefold():
        add(result, "errors", f"{label} artifact.sha256 is missing or mismatched")
    if renderer.get("input_sha256") != artifact_hash:
        add(result, "errors", f"{label} renderer.input_sha256 does not bind the delivered artifact")
    pages = data.get("pages")
    if not isinstance(pages, list) or not pages:
        add(result, "errors", f"{label} pages must be a nonempty ordered array")
        return
    artifact_page_count = data.get("artifact_page_count")
    if (
        not isinstance(artifact_page_count, int)
        or isinstance(artifact_page_count, bool)
        or artifact_page_count <= 0
        or artifact_page_count != len(pages)
    ):
        add(result, "errors", f"{label} artifact_page_count must equal the ordered pages length")
    if expected_pages is not None and artifact_page_count != expected_pages:
        add(
            result,
            "errors",
            f"{label} artifact_page_count {artifact_page_count!r} != expected {expected_pages}",
        )
    candidates: list[Path] = []
    for index, page in enumerate(pages, 1):
        if not isinstance(page, dict) or entry_page_number(page) != index:
            add(result, "errors", f"{label} page order mismatch at entry {index}")
            continue
        image_value = page.get("path") or page.get("image")
        expected_hash = page.get("sha256")
        if not isinstance(image_value, str):
            add(result, "errors", f"{label} page {index} has no image path")
            continue
        image = resolve_local_path(resolved.parent, image_value)
        if not image.is_file() or image.suffix.lower() not in IMAGE_EXTENSIONS:
            add(result, "errors", f"{label} page {index} image is missing or unsupported: {image}")
            continue
        if not isinstance(expected_hash, str) or expected_hash.casefold() != sha256_file(image).casefold():
            add(result, "errors", f"{label} page {index} SHA-256 is missing or mismatched")
        candidates.append(image)
    if expected_pages is not None and len(pages) != expected_pages:
        add(result, "errors", f"{label} pages {len(pages)} != expected {expected_pages}")
    hashes: list[str] = []
    for index, image in enumerate(candidates, 1):
        try:
            width, height, _, _ = decode_image_path(image)
        except Exception as exc:
            add(result, "errors", f"{label} page {index} cannot be fully decoded: {exc}")
            continue
        # Readability is orientation agnostic: slide renders are landscape,
        # while report/notes renders are commonly portrait (for example,
        # 1275x1650 from LibreOffice at 150 DPI).
        if min(width, height) < 720 or max(width, height) < 1280:
            add(result, "errors", f"{label} page {index} is below readable size: {width}x{height}")
        if Image is not None:
            try:
                with Image.open(image) as rendered:
                    gray = ImageOps.exif_transpose(rendered).convert("L").resize((64, 64))
                    extrema = gray.getextrema()
                    histogram = gray.histogram()
                    occupied_bins = sum(1 for count in histogram if count)
                    if (
                        not extrema
                        or extrema[1] - extrema[0] < 4
                        or occupied_bins < 4
                    ):
                        add(result, "errors", f"{label} page {index} appears blank or near-uniform")
            except Exception as exc:
                add(result, "errors", f"{label} page {index} blank-page inspection failed: {exc}")
        hashes.append(sha256_file(image))
    if len(set(hashes)) != len(hashes):
        add(result, "errors", f"{label} contains duplicate rendered pages")
    independent_pages = independently_render_artifact_pages(artifact, result, label)
    if independent_pages is None:
        add(
            result,
            "errors",
            f"{label} could not independently render the artifact for pixel reconciliation",
        )
    else:
        if len(independent_pages) != len(candidates):
            add(
                result,
                "errors",
                f"{label} independent render pages {len(independent_pages)} != evidence pages {len(candidates)}",
            )
        for index, (candidate, independent) in enumerate(zip(candidates, independent_pages), 1):
            try:
                _, _, candidate_pixels, candidate_phash = decode_image_path(candidate)
                _, _, independent_pixels, independent_phash = decode_image_path(independent)
            except Exception as exc:
                add(result, "errors", f"{label} page {index} independent pixel comparison failed: {exc}")
                continue
            if candidate_pixels != independent_pixels and hamming_distance(candidate_phash, independent_phash) > 8:
                add(result, "errors", f"{label} page {index} does not match an independent artifact render")
    result["metrics"][f"{label.replace(' ', '_')}_pages"] = len(candidates)


def independently_render_artifact_pages(
    artifact: Path,
    result: dict[str, Any],
    label: str,
) -> list[Path] | None:
    """Render DOCX/PPTX through a local office pipeline, independent of manifest claims."""
    try:
        return render_pages(artifact)
    except (OSError, subprocess.SubprocessError, ValueError, RuntimeError):
        return None


def validate_editable_visual_record(
    record: dict[str, Any],
    base: Path,
    label: str,
    result: dict[str, Any],
    *,
    require_pair: bool,
    page_index: int | None = None,
    approved_source: ImageRecord | None = None,
    independent_editable: Path | None = None,
) -> dict[str, Path]:
    required = ("source_render", "editable_render", "comparison") if require_pair else ("comparison",)
    resolved: dict[str, Path] = {}
    for key in required:
        value = record.get(key)
        if value is None:
            add(result, "errors", f"{label} requires {key} visual evidence")
            continue
        validate_bound_evidence_file(value, base, f"{label} {key}", result)
        path_value = value.get("path") or value.get("file") if isinstance(value, dict) else value
        if isinstance(path_value, str):
            candidate = resolve_local_path(base, path_value)
            if candidate.is_file() and candidate.suffix.lower() in IMAGE_EXTENSIONS:
                resolved[key] = candidate
    if require_pair:
        source = resolved.get("source_render")
        editable = resolved.get("editable_render")
        comparison = resolved.get("comparison")
        if source is not None:
            if approved_source is None:
                add(result, "errors", f"{label} cannot bind source_render without the approved slide source")
            else:
                try:
                    _, _, source_pixels, source_phash = decode_image_path(source)
                except Exception as exc:
                    add(result, "errors", f"{label} source_render comparison failed: {exc}")
                else:
                    if source_pixels != approved_source.pixel_sha256 and hamming_distance(
                        source_phash, approved_source.perceptual_hash
                    ) > 4:
                        add(result, "errors", f"{label} source_render does not match the approved slide image")
        if editable is not None:
            if independent_editable is None:
                add(result, "errors", f"{label} cannot bind editable_render to an independent final-PPTX render")
            else:
                try:
                    _, _, editable_pixels, editable_phash = decode_image_path(editable)
                    _, _, independent_pixels, independent_phash = decode_image_path(independent_editable)
                except Exception as exc:
                    add(result, "errors", f"{label} editable_render comparison failed: {exc}")
                else:
                    if editable_pixels != independent_pixels and hamming_distance(
                        editable_phash, independent_phash
                    ) > 8:
                        add(result, "errors", f"{label} editable_render does not match independent final-PPTX page {page_index}")
        if source is not None and editable is not None and source.resolve() == editable.resolve():
            add(result, "errors", f"{label} source_render and editable_render must be distinct evidence files")
        if comparison is not None:
            if comparison.resolve() in {
                candidate.resolve() for candidate in (source, editable) if candidate is not None
            }:
                add(result, "errors", f"{label} comparison must be distinct from source/editable renders")
            if image_appears_blank(comparison):
                add(result, "errors", f"{label} comparison appears blank or near-uniform")
            if source is not None and editable is not None:
                validate_contact_sheet_pair(
                    comparison,
                    source,
                    editable,
                    f"{label} comparison",
                    result,
                )
    return resolved


def validate_contact_sheet_contains(
    comparison: Path,
    required_images: Sequence[Path],
    label: str,
    result: dict[str, Any],
) -> None:
    """Require a contact sheet to visibly contain each bound source/render.

    It is deliberately independent of self-declared comparison metadata.  The
    common two-column comparison is checked first with a bounded grid search;
    no unbounded pixel-by-pixel sliding window is used.
    """
    if Image is None or ImageOps is None:
        add(result, "errors", f"{label} needs Pillow for contact-sheet pixel reconciliation")
        return
    try:
        with Image.open(comparison) as opened:
            canvas = ImageOps.exif_transpose(opened).convert("RGB")
        if not required_images:
            return
        grid_candidates: list[Any] = []
        count = len(required_images)
        grid_shapes: set[tuple[int, int]] = set()
        # Enumerate exact-count grids in both orientations (for example 4x2
        # and 2x4 for an 8-image original+revised contact sheet), plus common
        # near-square grids with at most one trailing empty cell.
        for columns in range(1, count + 1):
            rows = (count + columns - 1) // columns
            if columns * rows in {count, count + 1}:
                grid_shapes.add((columns, rows))
                grid_shapes.add((rows, columns))
        grid_shapes.update({(2, 1), (1, 2), (2, 2), (3, 2), (2, 3), (3, 3)})
        for columns, rows in sorted(grid_shapes):
            for row in range(rows):
                for column in range(columns):
                    grid_candidates.append(
                        canvas.crop(
                            (
                                column * canvas.width // columns,
                                row * canvas.height // rows,
                                (column + 1) * canvas.width // columns,
                                (row + 1) * canvas.height // rows,
                            )
                        )
                    )
        for required in required_images:
            with Image.open(required) as opened:
                needle = ImageOps.exif_transpose(opened).convert("RGB")
            _, _, needle_pixels, needle_phash = _decoded_image(needle)
            matched = False
            for candidate in grid_candidates:
                _, _, candidate_pixels, candidate_phash = _decoded_image(candidate)
                if candidate_pixels == needle_pixels or hamming_distance(candidate_phash, needle_phash) <= 10:
                    matched = True
                    break
            if not matched:
                add(result, "errors", f"{label} does not visibly contain bound image {required.name}")
    except Exception as exc:
        add(result, "errors", f"{label} contact-sheet reconciliation failed: {exc}")


def validate_contact_sheet_pair(
    comparison: Path,
    before: Path,
    after: Path,
    label: str,
    result: dict[str, Any],
) -> None:
    """Validate the standard side-by-side pair at a bounded set of layouts."""
    if Image is None or ImageOps is None:
        add(result, "errors", f"{label} needs Pillow for contact-sheet reconciliation")
        return
    try:
        with Image.open(comparison) as opened:
            canvas = ImageOps.exif_transpose(opened).convert("RGB")
        with Image.open(before) as opened:
            before_hash = _decoded_image(ImageOps.exif_transpose(opened).convert("RGB"))[3]
        with Image.open(after) as opened:
            after_hash = _decoded_image(ImageOps.exif_transpose(opened).convert("RGB"))[3]

        def phash(crop: Any) -> int:
            return _decoded_image(crop)[3]

        pairs: list[tuple[Any, Any]] = []
        # Full half-panels.
        pairs.append(
            (
                canvas.crop((0, 0, canvas.width // 2, canvas.height)),
                canvas.crop((canvas.width // 2, 0, canvas.width, canvas.height)),
            )
        )
        # Comparison sheets often add labels/margins and place 16:9 panels in
        # the vertical center; test bounded central bands as well.
        for top_fraction, bottom_fraction in ((0.10, 0.90), (0.20, 0.80), (0.25, 0.75), (0.30, 0.70)):
            top = round(canvas.height * top_fraction)
            bottom = round(canvas.height * bottom_fraction)
            pairs.append(
                (
                    canvas.crop((0, top, canvas.width // 2, bottom)),
                    canvas.crop((canvas.width // 2, top, canvas.width, bottom)),
                )
            )
        matched = any(
            hamming_distance(phash(left), before_hash) <= 10
            and hamming_distance(phash(right), after_hash) <= 10
            for left, right in pairs
        )
        if not matched:
            add(result, "errors", f"{label} does not visibly contain the bound before/after pair")
    except Exception as exc:
        add(result, "errors", f"{label} contact-sheet pair reconciliation failed: {exc}")


def validate_editppt_evidence(
    path: Path,
    result: dict[str, Any],
    editable_pptx: Path | None,
    approved_sources: list[ImageRecord] | None,
) -> int | None:
    resolved = path.expanduser().resolve()
    run_dir: Path | None = resolved if resolved.is_dir() else None
    supplied_data: Any = None
    if resolved.is_file():
        supplied_data = parse_json_file(resolved, "editable-validation", result)
        if supplied_data is None:
            return None
        if resolved.name in {"page_jobs.json", "deck_manifest.json", "run_summary.json"}:
            run_dir = resolved.parent.parent if resolved.parent.name == "final" else resolved.parent
    elif not resolved.is_dir():
        add(result, "errors", f"editable-validation missing: {resolved}")
        return None

    bound_pptx_for_render = editable_pptx if editable_pptx is not None and editable_pptx.is_file() else None
    independent_editable_pages = (
        independently_render_artifact_pages(bound_pptx_for_render, result, "editable-validation")
        if bound_pptx_for_render is not None
        else None
    )
    if bound_pptx_for_render is not None and independent_editable_pages is None:
        add(result, "errors", "editable-validation could not independently render the final PPTX")
    if approved_sources is not None and independent_editable_pages is not None and len(approved_sources) != len(independent_editable_pages):
        add(result, "errors", "editable-validation approved-source and independent-render page counts differ")

    if run_dir is None:
        if not isinstance(supplied_data, dict) or supplied_data.get("passed") is not True:
            add(result, "errors", "editable-validation JSON must have top-level passed: true")
            return None
        evidence_hash = supplied_data.get("pptx_sha256") or supplied_data.get("sha256")
        evidence_pptx = supplied_data.get("pptx") or supplied_data.get("output")
        bound_pptx = editable_pptx if editable_pptx is not None and editable_pptx.is_file() else None
        if bound_pptx is None and isinstance(evidence_pptx, str):
            candidate = resolve_local_path(resolved.parent, evidence_pptx)
            if candidate.is_file():
                bound_pptx = candidate
        if bound_pptx is None:
            add(result, "errors", "standalone editable-validation does not identify an existing PPTX")
        elif not isinstance(evidence_hash, str):
            add(result, "errors", "standalone editable-validation has no pptx_sha256 binding")
        elif evidence_hash.casefold() != sha256_file(bound_pptx).casefold():
            add(result, "errors", "standalone editable-validation belongs to a different PPTX (SHA-256 mismatch)")
        pages = supplied_data.get("pages")
        if not isinstance(pages, list) or not pages:
            add(result, "errors", "standalone editable-validation must include a nonempty pages[] array")
            return None
        count = len(pages)
        for index, page in enumerate(pages, 1):
            if not isinstance(page, dict) or page.get("passed") is not True:
                add(result, "errors", f"editable-validation page {index} is not passed")
                continue
            if entry_page_number(page) != index:
                add(result, "errors", f"editable-validation page order mismatch at entry {index}")
            validation_record = load_bound_validation_record(
                page.get("validation"),
                resolved.parent,
                f"editable-validation page {index}",
                result,
                page.get("validation_sha256"),
            )
            if validation_record is not None:
                page_validation_value = page.get("validation")
                if isinstance(page_validation_value, dict):
                    page_validation_value = page_validation_value.get("path") or page_validation_value.get("file")
                page_validation_base = (
                    resolve_local_path(resolved.parent, page_validation_value).parent
                    if isinstance(page_validation_value, str)
                    else resolved.parent
                )
                validate_editable_visual_record(
                    validation_record,
                    page_validation_base,
                    f"editable-validation page {index}",
                    result,
                    require_pair=True,
                    page_index=index,
                    approved_source=approved_sources[index - 1] if approved_sources and index <= len(approved_sources) else None,
                    independent_editable=(
                        independent_editable_pages[index - 1]
                        if independent_editable_pages and index <= len(independent_editable_pages)
                        else None
                    ),
                )
                record_page = validation_record.get("page") or validation_record.get("page_index")
                if record_page is not None:
                    try:
                        if int(record_page) != index:
                            add(result, "errors", f"editable-validation page {index} record maps to page {record_page}")
                    except (TypeError, ValueError):
                        add(result, "errors", f"editable-validation page {index} record has invalid page mapping {record_page!r}")
        final_evidence = supplied_data.get("final_validation")
        final_record = load_bound_validation_record(
            final_evidence,
            resolved.parent,
            "editable-validation final",
            result,
            supplied_data.get("final_validation_sha256"),
        )
        if final_record is not None:
            final_path_value = final_evidence if isinstance(final_evidence, str) else final_evidence.get("path") if isinstance(final_evidence, dict) else None
            final_base = resolve_local_path(resolved.parent, final_path_value).parent if isinstance(final_path_value, str) else resolved.parent
            validate_editable_visual_record(
                final_record,
                final_base,
                "editable-validation final",
                result,
                require_pair=False,
            )
            final_count = final_record.get("slides") or final_record.get("expected_pages")
            if isinstance(final_count, int) and final_count != count:
                add(result, "errors", f"standalone final validation slides {final_count} != pages {count}")
        return int(count) if isinstance(count, int) else None

    deck_path = run_dir / "deck_manifest.json"
    jobs_path = run_dir / "page_jobs.json"
    deck = parse_json_file(deck_path, "editppt deck_manifest", result)
    jobs = parse_json_file(jobs_path, "editppt page_jobs", result)
    if not isinstance(deck, dict) or not isinstance(jobs, dict):
        return None
    if deck.get("schema_version") != 1 or jobs.get("schema_version") != 1:
        add(result, "errors", "editppt deck_manifest and page_jobs require schema_version: 1")
    if jobs.get("run_status") not in {"complete", "accepted"}:
        add(result, "errors", f"editppt page_jobs run_status is not complete: {jobs.get('run_status')!r}")
    pages = jobs.get("pages")
    if not isinstance(pages, list) or not pages:
        add(result, "errors", "editppt page_jobs has no pages")
        return 0
    page_numbers: list[int] = []
    for index, page in enumerate(pages, 1):
        if not isinstance(page, dict):
            add(result, "errors", f"editppt page_jobs page {index} is not an object")
            continue
        page_number = entry_page_number(page) or index
        page_numbers.append(page_number)
        status = str(page.get("status", ""))
        if status not in {"accepted", "complete"}:
            add(result, "errors", f"editppt page {page_number} is not finalized: status={status!r}")
        if page.get("accepted") is not True:
            add(result, "errors", f"editppt page {page_number} is not accepted")
        page_result = page.get("result")
        if not isinstance(page_result, dict) or page_result.get("validation_passed") is not True:
            add(result, "errors", f"editppt page {page_number} has no recorded validation_passed: true")
        if isinstance(page_result, dict):
            outputs = page_result.get("outputs")
            hashes = page_result.get("hashes")
            if not isinstance(outputs, dict) or not outputs:
                add(result, "errors", f"editppt page {page_number} has no recorded outputs")
            elif not isinstance(hashes, dict):
                add(result, "errors", f"editppt page {page_number} has no recorded output hashes")
            else:
                for role, output_value in outputs.items():
                    if not isinstance(output_value, str):
                        add(result, "errors", f"editppt page {page_number} output {role} has no path")
                        continue
                    output_path = resolve_local_path(run_dir, output_value)
                    if not output_path.is_file():
                        add(result, "errors", f"editppt page {page_number} output {role} is missing: {output_path}")
                        continue
                    expected_output_hash = hashes.get(role)
                    if not isinstance(expected_output_hash, str):
                        add(result, "errors", f"editppt page {page_number} output {role} has no SHA-256")
                    elif expected_output_hash.casefold() != sha256_file(output_path).casefold():
                        add(result, "errors", f"editppt page {page_number} output {role} SHA-256 mismatch")
        validation_value = page.get("validation")
        if not validation_value and isinstance(page_result, dict):
            outputs = page_result.get("outputs")
            if isinstance(outputs, dict):
                validation_value = outputs.get("validation")
        if not isinstance(validation_value, str):
            add(result, "errors", f"editppt page {page_number} has no validation path")
        else:
            validation_path = resolve_local_path(run_dir, validation_value)
            validation = parse_json_file(validation_path, f"editppt page {page_number} validation", result)
            if not isinstance(validation, dict) or validation.get("passed") is not True:
                add(result, "errors", f"editppt page {page_number} validation did not pass")
            elif isinstance(validation, dict):
                validate_editable_visual_record(
                    validation,
                    validation_path.parent,
                    f"editppt page {page_number} validation",
                    result,
                    require_pair=True,
                    page_index=page_number,
                    approved_source=(
                        approved_sources[page_number - 1]
                        if approved_sources and page_number <= len(approved_sources)
                        else None
                    ),
                    independent_editable=(
                        independent_editable_pages[page_number - 1]
                        if independent_editable_pages and page_number <= len(independent_editable_pages)
                        else None
                    ),
                )
    if page_numbers != list(range(1, len(pages) + 1)):
        add(result, "errors", f"editppt page order must be 1..{len(pages)}: {page_numbers}")

    deck_count = deck.get("page_count")
    if isinstance(deck_count, int) and deck_count != len(pages):
        add(result, "errors", f"editppt deck_manifest page_count {deck_count} != page_jobs pages {len(pages)}")
    deck_pages = deck.get("pages")
    if isinstance(deck_pages, list) and len(deck_pages) != len(pages):
        add(result, "errors", f"editppt deck_manifest pages {len(deck_pages)} != page_jobs pages {len(pages)}")

    final_validation = first_existing(
        [
            run_dir / "final" / "validation.json",
            run_dir / "validation.json",
            resolve_local_path(run_dir, str(deck.get("validation"))) if deck.get("validation") else run_dir / "__missing__",
        ]
    )
    if final_validation is None:
        add(result, "errors", "editppt final validation.json is missing")
    else:
        validation = parse_json_file(final_validation, "editppt final validation", result)
        if not isinstance(validation, dict) or validation.get("passed") is not True:
            add(result, "errors", "editppt final validation did not pass")
        elif isinstance(validation, dict):
            validate_editable_visual_record(
                validation,
                final_validation.parent,
                "editppt final validation",
                result,
                require_pair=False,
            )
            if isinstance(validation.get("slides"), int) and validation["slides"] != len(pages):
                add(result, "errors", f"editppt final validation slides {validation['slides']} != pages {len(pages)}")

    output_value = deck.get("output")
    output_candidates: list[Path] = []
    if isinstance(output_value, str):
        output_candidates.extend([resolve_local_path(run_dir, output_value), run_dir / "final" / Path(output_value).name])
    output_path = first_existing(output_candidates)
    if output_path is None:
        add(result, "errors", "editppt finalized output PPTX is missing")
    elif editable_pptx is not None and editable_pptx.is_file():
        if sha256_file(output_path) != sha256_file(editable_pptx):
            add(result, "errors", "editable-validation evidence belongs to a different PPTX (SHA-256 mismatch)")
    result["metrics"]["editable_validation_pages"] = len(pages)
    return len(pages)


def validate_text_outline(path: Path, result: dict[str, Any]) -> str | None:
    if not path.is_file():
        add(result, "errors", f"outline missing or not a file: {path}")
        return None
    if path.stat().st_size == 0:
        add(result, "errors", f"outline is empty: {path}")
        return None
    if path.suffix.lower() in {".md", ".txt"}:
        try:
            text = path.read_text(encoding="utf-8")
        except (OSError, UnicodeError) as exc:
            add(result, "errors", f"outline cannot be read as UTF-8 text: {exc}")
            return None
        if not normalize_text(text):
            add(result, "errors", "outline has no visible text")
        hits = placeholder_hits(text.splitlines())
        if hits:
            add(result, "errors", f"outline contains placeholder markers {hits}")
        return text
    elif path.suffix.lower() == ".json":
        data = parse_json_file(path, "outline", result)
        if data is not None:
            entries = ledger_entries(data, "outline", result)
            if entries is not None:
                validate_ledger_contract(data, entries, "slide-copy ledger", result)
                return "\n".join(
                    "\n".join(entry_text(entry, SLIDE_TEXT_FIELDS)) for entry in entries
                )
    else:
        add(result, "errors", f"outline format {path.suffix or '(none)'} is unsupported for deterministic reconciliation")
    return None


def reconcile_outline_to_slide_copy(
    outline_text: str | None,
    slide_copy: list[dict[str, Any]] | None,
    result: dict[str, Any],
) -> None:
    if slide_copy is None:
        return
    if not normalize_text(outline_text):
        add(result, "errors", "outline text could not be reconciled to the slide-copy ledger")
        return
    for page, entry in enumerate(slide_copy, 1):
        expected = entry_text(
            entry,
            ("title", "purpose", "body_copy", "body", "copy", "critical_fields", "footer", "source_footer", "speaker_takeaway"),
        )
        missing = [value for value in expected if not text_contains(outline_text or "", value)]
        if missing:
            add(result, "errors", f"outline is missing slide-copy page {page} content: {missing}")


def required_mapping(data: dict[str, Any], key: str, label: str, result: dict[str, Any]) -> dict[str, Any] | None:
    value = data.get(key)
    if not isinstance(value, dict) or not value:
        add(result, "errors", f"{label} requires a nonempty {key} object")
        return None
    return value


def validate_visual_generation_provenance(
    data: Any,
    label: str,
    data_handling: str | None,
    result: dict[str, Any],
) -> None:
    if data_handling != "local-only":
        return
    provenance = data.get("generation_provenance") if isinstance(data, dict) else None
    if not isinstance(provenance, dict):
        add(result, "errors", f"{label} requires generation_provenance under local-only handling")
        return
    if provenance.get("data_handling") != "local-only":
        add(result, "errors", f"{label} generation_provenance.data_handling must be local-only")
    backend = provenance.get("backend")
    if backend not in {"local-native", "local-pillow", "user-provided", "disabled"}:
        add(result, "errors", f"{label} generation_provenance.backend is not an approved local backend")
    if provenance.get("remote_generation") is not False:
        add(result, "errors", f"{label} generation_provenance.remote_generation must be false")
    if provenance.get("uploaded") is not False:
        add(result, "errors", f"{label} generation_provenance.uploaded must be false")


def validate_style_options(
    path: Path,
    result: dict[str, Any],
    require_selection: bool = True,
    data_handling: str | None = None,
) -> None:
    resolved = path.expanduser().resolve()
    data = parse_json_file(resolved, "style-options", result)
    if not isinstance(data, dict):
        return
    if data.get("schema_version") != 1:
        add(result, "errors", "style-options requires schema_version: 1")
    validate_visual_generation_provenance(data, "style-options", data_handling, result)
    comparison_id = data.get("comparison_content_id")
    if not isinstance(comparison_id, str) or not comparison_id.strip():
        add(result, "errors", "style-options requires comparison_content_id")
    candidates = data.get("candidates")
    if not isinstance(candidates, list) or not 2 <= len(candidates) <= 3:
        add(result, "errors", "style-options must contain 2 or 3 candidates")
        return
    seen_ids: set[str] = set()
    valid_ids: set[str] = set()
    image_hashes: list[str] = []
    for index, candidate in enumerate(candidates, 1):
        if not isinstance(candidate, dict):
            add(result, "errors", f"style-options candidate {index} is not an object")
            continue
        candidate_id = candidate.get("id")
        if not isinstance(candidate_id, str) or not candidate_id.strip():
            add(result, "errors", f"style-options candidate {index} has no id")
            continue
        if candidate_id in seen_ids:
            add(result, "errors", f"style-options repeats candidate id {candidate_id!r}")
        seen_ids.add(candidate_id)
        image_value = candidate.get("image") or candidate.get("path")
        if not isinstance(image_value, str):
            add(result, "errors", f"style-options candidate {candidate_id} has no image path")
            continue
        if not Path(image_value).expanduser().is_absolute():
            add(result, "errors", f"style-options candidate {candidate_id} image path must be absolute")
        image_path = resolve_local_path(resolved.parent, image_value)
        if not image_path.is_file():
            add(result, "errors", f"style-options candidate {candidate_id} image is missing: {image_path}")
            continue
        try:
            width, height, _, _ = decode_image_path(image_path)
        except Exception as exc:
            add(result, "errors", f"style-options candidate {candidate_id} image cannot be fully decoded: {exc}")
            continue
        if width * 9 != height * 16:
            add(result, "errors", f"style-options candidate {candidate_id} is not exact 16:9: {width}x{height}")
        expected_width = candidate.get("width")
        expected_height = candidate.get("height")
        try:
            if expected_width is None or int(expected_width) != width:
                add(result, "errors", f"style-options candidate {candidate_id} width is missing or mismatched")
            if expected_height is None or int(expected_height) != height:
                add(result, "errors", f"style-options candidate {candidate_id} height is missing or mismatched")
        except (TypeError, ValueError):
            add(result, "errors", f"style-options candidate {candidate_id} dimensions are not numeric")
        expected_hash = candidate.get("sha256")
        actual_hash = sha256_file(image_path)
        image_hashes.append(actual_hash)
        if not isinstance(expected_hash, str) or expected_hash.casefold() != actual_hash.casefold():
            add(result, "errors", f"style-options candidate {candidate_id} SHA-256 is missing or mismatched")
        if candidate.get("hard_gate") != "passed":
            add(result, "errors", f"style-options candidate {candidate_id} hard_gate is not passed")
        if not isinstance(candidate.get("parameters"), dict) or not candidate["parameters"]:
            add(result, "errors", f"style-options candidate {candidate_id} has no visual parameters")
        prompt_value = candidate.get("prompt_file")
        if not isinstance(prompt_value, str) or not Path(prompt_value).expanduser().is_absolute():
            add(result, "errors", f"style-options candidate {candidate_id} prompt_file path must be absolute")
        elif not resolve_local_path(resolved.parent, prompt_value).is_file():
            add(result, "errors", f"style-options candidate {candidate_id} prompt_file is missing")
        valid_ids.add(candidate_id)
    if len(set(image_hashes)) != len(image_hashes):
        add(result, "errors", "style-options candidates are not visually distinct files (duplicate SHA-256)")
    selected = data.get("selected_id")
    if selected is None and not require_selection:
        if data.get("selected_reason") not in {None, ""}:
            add(result, "errors", "style-options has selected_reason but no selected_id")
        add(result, "warnings", "style-options candidates are accepted but no style is selected; downstream visual work remains paused")
    else:
        if not isinstance(selected, str) or selected not in valid_ids:
            add(result, "errors", f"style-options selected_id {selected!r} does not identify a candidate")
        if not normalize_text(data.get("selected_reason")):
            add(result, "errors", "style-options requires selected_reason")
    result["metrics"]["style_options_candidates"] = len(candidates)
    result["metrics"]["style_options_selected_id"] = selected


def validate_visual_system(
    path: Path,
    result: dict[str, Any],
    style_options_path: Path | None = None,
    data_handling: str | None = None,
) -> None:
    resolved = path.expanduser().resolve()
    data = parse_json_file(resolved, "visual-system", result)
    if not isinstance(data, dict):
        return
    if data.get("schema_version") != 1:
        add(result, "errors", "visual-system requires schema_version: 1")
    validate_visual_generation_provenance(data, "visual-system", data_handling, result)
    source_style = data.get("source_style_options")
    source_style_hash = data.get("source_style_options_sha256")
    selected_id = data.get("selected_id")
    if style_options_path is not None:
        expected_style = style_options_path.expanduser().resolve()
        if not isinstance(source_style, str) or Path(source_style).expanduser().resolve() != expected_style:
            add(result, "errors", "visual-system source_style_options does not match --style-options")
        if not expected_style.is_file():
            add(result, "errors", f"style-options linkage file is missing: {expected_style}")
        elif not isinstance(source_style_hash, str) or source_style_hash.casefold() != sha256_file(expected_style).casefold():
            add(result, "errors", "visual-system source_style_options_sha256 is missing or mismatched")
        style_data = parse_json_file(expected_style, "style-options linkage", result)
        if isinstance(style_data, dict) and selected_id != style_data.get("selected_id"):
            add(result, "errors", "visual-system selected_id does not match style-options selected_id")
    elif source_style is not None or source_style_hash is not None or selected_id is not None:
        add(result, "warnings", "visual-system style linkage could not be reconciled without --style-options")
    if not normalize_text(data.get("name")):
        add(result, "errors", "visual-system requires name")
    palette = required_mapping(data, "palette", "visual-system", result)
    typography = required_mapping(data, "typography", "visual-system", result)
    geometry = required_mapping(data, "geometry", "visual-system", result)
    required_mapping(data, "imagery", "visual-system", result)
    required_mapping(data, "charts", "visual-system", result)
    required_mapping(data, "icons", "visual-system", result)
    required_mapping(data, "footer", "visual-system", result)
    if palette is not None:
        for key in ("background", "primary", "secondary", "accent"):
            value = palette.get(key)
            if not isinstance(value, str) or not re.fullmatch(r"#[0-9A-Fa-f]{6}", value):
                add(result, "errors", f"visual-system palette.{key} must be a six-digit hex color")
    if typography is not None:
        for key in ("title_font", "body_font"):
            if not normalize_text(typography.get(key)):
                add(result, "errors", f"visual-system typography requires {key}")
        for key in ("title_weight", "body_weight"):
            value = typography.get(key)
            if not isinstance(value, int) or isinstance(value, bool) or not 100 <= value <= 1000:
                add(result, "errors", f"visual-system typography requires numeric {key} between 100 and 1000")
    if geometry is not None:
        canvas = geometry.get("canvas")
        if not (
            isinstance(canvas, list)
            and len(canvas) == 2
            and all(isinstance(value, int) and not isinstance(value, bool) and value > 0 for value in canvas)
            and canvas[0] * 9 == canvas[1] * 16
        ):
            add(result, "errors", "visual-system geometry.canvas must be an exact 16:9 [width,height]")
        for key in ("margin", "grid_columns", "gutter"):
            if not isinstance(geometry.get(key), (int, float)) or geometry[key] <= 0:
                add(result, "errors", f"visual-system geometry requires positive {key}")
    for key, required_key in (("imagery", "treatment"), ("charts", "style"), ("icons", "style")):
        value = data.get(key)
        if isinstance(value, dict) and not normalize_text(value.get(required_key)):
            add(result, "errors", f"visual-system {key} requires {required_key}")
    footer = data.get("footer")
    if isinstance(footer, dict):
        if not normalize_text(footer.get("position")):
            add(result, "errors", "visual-system footer requires position")
        minimum_size = footer.get("minimum_size_pt")
        if not isinstance(minimum_size, (int, float)) or isinstance(minimum_size, bool) or minimum_size <= 0:
            add(result, "errors", "visual-system footer requires positive minimum_size_pt")
    layouts = data.get("layouts")
    layout_names = [normalize_text(item) for item in layouts] if isinstance(layouts, list) else []
    expected_layouts = {"cover", "section", "narrative", "comparison", "data-chart", "conclusion"}
    if len(layout_names) != 6 or set(layout_names) != expected_layouts:
        add(result, "errors", f"visual-system layouts must be exactly {sorted(expected_layouts)}")
    result["metrics"]["visual_system_layouts"] = len(layouts) if isinstance(layouts, list) else 0


def validate_template_layouts(
    directory: Path,
    result: dict[str, Any],
    min_width: int,
    min_height: int,
    composer_path: Path | None,
) -> None:
    resolved = directory.expanduser().resolve()
    if not resolved.is_dir():
        add(result, "errors", f"template layout directory missing: {resolved}")
        return
    images = [path for path in resolved.iterdir() if path.is_file() and path.suffix.lower() in IMAGE_EXTENSIONS]
    if len(images) != 6:
        add(result, "errors", f"template layout directory must contain exactly six PNG/JPEG images, found {len(images)}")
    if composer_path is not None:
        # Composer inputs are the explicit semantic/order authority. Numeric
        # filenames are required only when no such manifest exists.
        ordered = [(image.absolute(), index) for index, image in enumerate(sorted(images, key=lambda item: item.name), 1)]
    else:
        ordered = order_slide_images(images, None, result) if images else []
    dimensions: list[list[int]] = []
    layout_hashes: dict[Path, str] = {}
    for image, _ in ordered:
        try:
            width, height, _, _ = decode_image_path(image)
        except Exception as exc:
            add(result, "errors", f"template layout image {image.name} cannot be fully decoded: {exc}")
            continue
        dimensions.append([width, height])
        if width * 9 != height * 16:
            add(result, "errors", f"template layout image is not exact 16:9: {image.name} ({width}x{height})")
        if width < min_width or height < min_height:
            add(result, "errors", f"template layout image is below {min_width}x{min_height}: {image.name}")
        layout_hashes[image.absolute()] = sha256_file(image)
    if len(set(layout_hashes.values())) != len(layout_hashes):
        duplicates = sorted(
            path.name for path, digest in layout_hashes.items()
            if sum(other == digest for other in layout_hashes.values()) > 1
        )
        add(result, "errors", f"template layout images must be six distinct files; duplicate content: {duplicates}")

    if composer_path is None:
        add(result, "warnings", "template layout semantics are unverified without --template-composer inputs")
    else:
        composer_resolved = composer_path.expanduser().resolve()
        composer = parse_json_file(composer_resolved, "template composer", result)
        if isinstance(composer, dict):
            if composer.get("schema_version") != 1:
                add(result, "errors", "template composer requires schema_version: 1")
            if composer.get("tool") != "ppt-gen-compose-template-overview":
                add(result, "errors", "template composer tool must be 'ppt-gen-compose-template-overview'")
            if composer.get("status") != "passed":
                add(result, "errors", "template composer status must be 'passed'")
            if composer.get("geometry_source") != "local-pillow":
                add(result, "errors", "template composer geometry_source must be 'local-pillow'")
        inputs = composer.get("inputs") if isinstance(composer, dict) else None
        if not isinstance(inputs, list) or len(inputs) != 6:
            add(result, "errors", "template composer inputs must contain exactly six layout entries")
        else:
            composer_paths: set[Path] = set()
            semantic_labels: list[str] = []
            for index, entry in enumerate(inputs, 1):
                if not isinstance(entry, dict):
                    add(result, "errors", f"template composer input {index} is not an object")
                    continue
                path_value = entry.get("path") or entry.get("image") or entry.get("file")
                if not isinstance(path_value, str):
                    add(result, "errors", f"template composer input {index} has no path")
                    continue
                input_path = resolve_local_path(composer_resolved.parent, path_value)
                input_identity = Path(path_value).expanduser()
                if not input_identity.is_absolute():
                    input_identity = composer_resolved.parent / input_identity
                input_identity = input_identity.absolute()
                composer_paths.add(input_identity)
                actual_hash = layout_hashes.get(input_identity)
                if actual_hash is None:
                    add(result, "errors", f"template composer input {index} is not a template-layout file: {input_path}")
                expected_hash = entry.get("sha256")
                if actual_hash is not None and (
                    not isinstance(expected_hash, str) or expected_hash.casefold() != actual_hash.casefold()
                ):
                    add(result, "errors", f"template composer input {index} SHA-256 is missing or mismatched")
                expected_width = entry.get("width")
                expected_height = entry.get("height")
                try:
                    width, height, _, _ = decode_image_path(input_path)
                    if int(expected_width) != width or int(expected_height) != height:
                        add(result, "errors", f"template composer input {index} dimensions mismatch")
                    if entry.get("exact_16_9") is not True or width * 9 != height * 16:
                        add(result, "errors", f"template composer input {index} is not recorded as exact 16:9")
                except (TypeError, ValueError, OSError, UnidentifiedImageError) as exc:
                    add(result, "errors", f"template composer input {index} has invalid image metadata: {exc}")
                label_value = entry.get("layout_id") or entry.get("label") or entry.get("name")
                semantic = normalize_layout_semantic(label_value)
                if semantic is None:
                    add(result, "errors", f"template composer input {index} has unknown layout label {label_value!r}")
                else:
                    semantic_labels.append(semantic)
            if composer_paths != set(layout_hashes):
                missing = sorted(path.name for path in set(layout_hashes) - composer_paths)
                extras = sorted(str(path) for path in composer_paths - set(layout_hashes))
                add(result, "errors", f"template composer inputs do not exactly match layout directory: missing={missing}, extra={extras}")
            required_semantics = {"cover", "section", "narrative", "comparison", "data-chart", "conclusion"}
            if len(semantic_labels) != 6 or set(semantic_labels) != required_semantics:
                add(result, "errors", f"template composer layout labels must cover exactly {sorted(required_semantics)}")
    result["metrics"]["template_layouts"] = len(images)
    result["metrics"]["template_layout_dimensions"] = dimensions


def normalize_layout_semantic(value: Any) -> str | None:
    text = compact_text(value).replace("_", "-")
    aliases = {
        "cover": {"cover", "封面", "封面页"},
        "section": {"section", "section-divider", "章节", "章节页", "分节", "过渡页"},
        "narrative": {"narrative", "content", "叙事", "叙事页", "叙事内容", "内容", "内容页"},
        "comparison": {"comparison", "compare", "对比", "对比页"},
        "data-chart": {"data-chart", "datachart", "chart", "数据", "数据图表", "数据页", "图表", "图表页"},
        "conclusion": {"conclusion", "closing", "summary", "结论", "结论页", "总结", "总结页"},
    }
    for semantic, names in aliases.items():
        if text in names:
            return semantic
    return None


def dimension_pair(value: Any) -> tuple[int, int] | None:
    if isinstance(value, (list, tuple)) and len(value) == 2:
        try:
            return int(value[0]), int(value[1])
        except (TypeError, ValueError):
            return None
    if isinstance(value, dict):
        try:
            return int(value.get("width")), int(value.get("height"))
        except (TypeError, ValueError):
            return None
    return None


def validate_template_overview(
    path: Path,
    composer_path: Path | None,
    visual_system_path: Path | None,
    result: dict[str, Any],
) -> None:
    resolved = path.expanduser().resolve()
    if not resolved.is_file():
        add(result, "errors", f"template overview image missing: {resolved}")
        return
    if resolved.suffix.lower() != ".png":
        add(result, "errors", f"template overview must be PNG: {resolved.name}")
    try:
        width, height, _, _ = decode_image_path(resolved)
    except Exception as exc:
        add(result, "errors", f"template overview image cannot be fully decoded: {exc}")
        return
    if (width, height) != (1920, 1080):
        add(result, "errors", f"template overview must be exactly 1920x1080: {width}x{height}")
    result["metrics"]["template_overview_size"] = [width, height]
    if composer_path is None:
        return
    composer = parse_json_file(composer_path.expanduser().resolve(), "template composer", result)
    if not isinstance(composer, dict):
        return
    if composer.get("schema_version") != 1:
        add(result, "errors", "template composer requires schema_version: 1")
    if composer.get("tool") != "ppt-gen-compose-template-overview":
        add(result, "errors", "template composer tool must be 'ppt-gen-compose-template-overview'")
    if composer.get("status") != "passed":
        add(result, "errors", "template composer status must be 'passed'")
    if composer.get("geometry_source") != "local-pillow":
        add(result, "errors", "template composer geometry_source must be 'local-pillow'")
    composer_visual = composer.get("visual_system") if isinstance(composer.get("visual_system"), dict) else {}
    if visual_system_path is not None:
        expected_visual = visual_system_path.expanduser().resolve()
        visual_value = composer_visual.get("path")
        if not isinstance(visual_value, str) or resolve_local_path(
            composer_path.expanduser().resolve().parent, visual_value
        ) != expected_visual:
            add(result, "errors", "template composer visual_system.path does not match --visual-system")
        visual_hash = composer_visual.get("sha256")
        if not expected_visual.is_file():
            add(result, "errors", f"template visual-system is missing: {expected_visual}")
        elif not isinstance(visual_hash, str) or visual_hash.casefold() != sha256_file(expected_visual).casefold():
            add(result, "errors", "template composer visual_system SHA-256 is missing or mismatched")
    output = composer.get("output") if isinstance(composer.get("output"), dict) else {}
    thumbnail = dimension_pair(output.get("thumbnail_size") or composer.get("thumbnail_size"))
    if thumbnail is None or thumbnail[0] <= 0 or thumbnail[1] <= 0 or thumbnail[0] * 9 != thumbnail[1] * 16:
        add(result, "errors", "template composer thumbnail_size must be exact 16:9")
    output_path_value = output.get("path") or composer.get("overview_path")
    if not isinstance(output_path_value, str) or resolve_local_path(composer_path.expanduser().resolve().parent, output_path_value) != resolved:
        add(result, "errors", "template composer output.path does not match --template-overview")
    if output.get("width") != 1920 or output.get("height") != 1080 or output.get("layout") != "3x2":
        add(result, "errors", "template composer output must record 1920x1080 and layout='3x2'")
    overview_hash = output.get("sha256") or composer.get("overview_sha256") or composer.get("sha256")
    if not isinstance(overview_hash, str) or overview_hash.casefold() != sha256_file(resolved).casefold():
        add(result, "errors", "template composer overview SHA-256 is missing or mismatched")


def validate_source_manifest(path: Path, result: dict[str, Any]) -> dict[str, Any] | None:
    data = parse_json_file(path, "source-manifest", result)
    if not isinstance(data, dict):
        return None
    if data.get("schema_version") != 1:
        add(result, "errors", "source-manifest requires schema_version: 1")
    manifest_errors = data.get("errors", [])
    manifest_warnings = data.get("warnings", [])
    if data.get("status") == "failed" or (isinstance(manifest_errors, list) and manifest_errors):
        add(result, "errors", "source-manifest records failed adaptation or hard errors")
    if manifest_errors is not None and not isinstance(manifest_errors, list):
        add(result, "errors", "source-manifest errors must be an array")
    if manifest_warnings is not None and not isinstance(manifest_warnings, list):
        add(result, "errors", "source-manifest warnings must be an array")
        manifest_warnings = []
    for warning in manifest_warnings or []:
        if normalize_text(warning):
            add(result, "warnings", f"source-manifest: {normalize_text(warning)}")
    if data.get("status") == "passed_with_warnings" and not manifest_warnings:
        add(result, "errors", "source-manifest status is passed_with_warnings but warnings are empty")
    if data.get("source_policy") not in {"faithful", "verify-update"}:
        add(result, "errors", "source-manifest source_policy must be 'faithful' or 'verify-update'")
    if data.get("data_handling") not in {"standard", "local-only"}:
        add(result, "errors", "source-manifest data_handling must be 'standard' or 'local-only'")
    if data.get("ocr_policy") != "offline":
        add(result, "errors", "source-manifest ocr_policy must be offline")
    if data.get("external_ocr_allowed") is not False:
        add(result, "errors", "source-manifest external_ocr_allowed must be false")
    if data.get("ask_for_ocr_token") is not False:
        add(result, "errors", "source-manifest ask_for_ocr_token must be false")
    sources = data.get("sources")
    if not isinstance(sources, list) or not sources:
        add(result, "errors", "source-manifest must contain a nonempty sources array")
        return data
    primary_source_id = data.get("primary_source_id")
    if not isinstance(primary_source_id, str) or not primary_source_id:
        add(result, "errors", "source-manifest requires primary_source_id")
    seen_ids: set[str] = set()
    for index, source in enumerate(sources, 1):
        if not isinstance(source, dict):
            add(result, "errors", f"source-manifest source {index} is not an object")
            continue
        source_id = source.get("id")
        if not isinstance(source_id, str) or not source_id:
            add(result, "errors", f"source-manifest source {index} has no id")
        elif source_id in seen_ids:
            add(result, "errors", f"source-manifest repeats source id {source_id!r}")
        else:
            seen_ids.add(source_id)
        path_value = source.get("path")
        if not isinstance(path_value, str) or not Path(path_value).expanduser().is_absolute():
            add(result, "errors", f"source-manifest source {index} path must be absolute")
            continue
        source_path = Path(path_value).expanduser().resolve()
        if not source_path.exists() or not (source_path.is_file() or source_path.is_dir()):
            add(result, "errors", f"source-manifest source {index} is missing: {source_path}")
            continue
        expected_hash = source.get("sha256")
        try:
            try:
                actual_hash = (
                    sha256_file(source_path)
                    if source_path.is_file()
                    else source_adapter_directory_hash(source_path)
                )
            except (OSError, ValueError) as exc:
                add(result, "errors", f"source-manifest source {index} cannot be safely hashed: {exc}")
                continue
        except OSError as exc:
            add(result, "errors", f"source-manifest source {index} cannot be hashed: {exc}")
            continue
        if not isinstance(expected_hash, str) or expected_hash.casefold() != actual_hash.casefold():
            add(result, "errors", f"source-manifest source {index} SHA-256 is missing or mismatched")
    if isinstance(primary_source_id, str) and primary_source_id not in seen_ids:
        add(result, "errors", "source-manifest primary_source_id does not identify a source")
    return data


def validate_claim_ledger(
    path: Path,
    result: dict[str, Any],
    source_manifest: dict[str, Any] | None,
) -> dict[str, Any] | None:
    """Validate factual authority and fail closed on unresolved downstream claims."""
    data = parse_json_file(path.expanduser().resolve(), "claim-ledger", result)
    if not isinstance(data, dict):
        return None
    if data.get("status") == "draft":
        add(result, "errors", "claim-ledger is still a draft; extract and verify claims before final QA")
    if contains_draft_sentinel(data):
        add(result, "errors", "claim-ledger still contains adapter draft sentinel text")
    if data.get("schema_version") != 1:
        add(result, "errors", "claim-ledger requires schema_version: 1")
    cutoff = data.get("source_cutoff")
    if not isinstance(cutoff, str) or not re.fullmatch(r"\d{4}-\d{2}-\d{2}", cutoff):
        add(result, "errors", "claim-ledger source_cutoff must be an ISO date (YYYY-MM-DD)")
    claims = data.get("claims")
    if not isinstance(claims, list):
        add(result, "errors", "claim-ledger claims must be an array")
        return data
    if not claims and not normalize_text(data.get("not_applicable_reason")):
        add(result, "errors", "claim-ledger requires claims or a nonempty not_applicable_reason")

    known_sources = {
        source.get("id")
        for source in (source_manifest or {}).get("sources", [])
        if isinstance(source, dict) and isinstance(source.get("id"), str)
    }
    seen_ids: set[str] = set()
    for index, claim in enumerate(claims, 1):
        if not isinstance(claim, dict):
            add(result, "errors", f"claim-ledger claim {index} is not an object")
            continue
        claim_id = claim.get("id")
        if not isinstance(claim_id, str) or not claim_id.strip():
            add(result, "errors", f"claim-ledger claim {index} requires id")
        elif claim_id in seen_ids:
            add(result, "errors", f"claim-ledger repeats id {claim_id!r}")
        else:
            seen_ids.add(claim_id)
        if not normalize_text(claim.get("statement")):
            add(result, "errors", f"claim-ledger claim {index} requires statement")
        kind = claim.get("kind")
        if kind not in CLAIM_KINDS:
            add(result, "errors", f"claim-ledger claim {index} kind must be one of {sorted(CLAIM_KINDS)}")
        verification = claim.get("verification")
        if verification not in CLAIM_VERIFICATIONS:
            add(
                result,
                "errors",
                f"claim-ledger claim {index} verification must be one of {sorted(CLAIM_VERIFICATIONS)}",
            )
        elif verification in {"pending", "unsupported"}:
            add(result, "errors", f"claim-ledger claim {claim_id or index} is unresolved ({verification})")
        source_ids = claim.get("source_ids")
        if not isinstance(source_ids, list) or not source_ids or not all(
            isinstance(value, str) and value for value in source_ids
        ):
            add(result, "errors", f"claim-ledger claim {index} source_ids must be a nonempty string array")
        elif source_manifest is not None:
            unknown = sorted(set(source_ids) - known_sources)
            if unknown:
                add(result, "errors", f"claim-ledger claim {index} references unknown source IDs: {unknown}")
        if kind in {"fact", "analysis"} and not normalize_text(claim.get("source_locator")):
            add(result, "errors", f"claim-ledger claim {index} requires source_locator")
        used_on_slides = claim.get("used_on_slides", [])
        if not isinstance(used_on_slides, list) or not all(
            isinstance(page, int) and not isinstance(page, bool) and page > 0 for page in used_on_slides
        ):
            add(result, "errors", f"claim-ledger claim {index} used_on_slides must be a positive-integer array")
    result["metrics"]["claims"] = len(claims)
    result["metrics"]["claim_ids"] = sorted(seen_ids)
    return data


def validate_claim_references(
    claim_data: dict[str, Any] | None,
    slide_copy: list[dict[str, Any]] | None,
    notes_ledger: list[dict[str, Any]] | None,
    result: dict[str, Any],
) -> None:
    if claim_data is None:
        return
    claims = [claim for claim in claim_data.get("claims", []) if isinstance(claim, dict)]
    known = {
        claim.get("id")
        for claim in claims
        if isinstance(claim.get("id"), str)
    }
    claims_by_id = {
        claim["id"]: claim
        for claim in claims
        if isinstance(claim.get("id"), str)
    }

    def evidence_for_page(claim: dict[str, Any], field: str, page: int) -> list[str]:
        value = claim.get(field)
        selected: Any = None
        if isinstance(value, dict):
            selected = value.get(str(page), value.get(page))
        elif isinstance(value, list):
            matching = [
                item for item in value
                if isinstance(item, dict) and entry_page_number(item) == page
            ]
            if matching:
                selected = matching
            elif all(not isinstance(item, dict) for item in value):
                used = claim.get("used_on_slides")
                if isinstance(used, list) and used == [page]:
                    selected = value
        return flatten_text_values(selected) if selected is not None else []
    for label, entries in (("slide-copy ledger", slide_copy), ("speaker-notes ledger", notes_ledger)):
        if entries is None:
            continue
        for page, entry in enumerate(entries, 1):
            claim_ids = entry.get("claim_ids")
            if claim_ids is None:
                continue
            if not isinstance(claim_ids, list) or not all(isinstance(value, str) for value in claim_ids):
                add(result, "errors", f"{label} page {page} claim_ids must be a string array")
                continue
            unknown = sorted(set(claim_ids) - known)
            if unknown:
                add(result, "errors", f"{label} page {page} references unknown claim IDs: {unknown}")

    if slide_copy is not None:
        ledger_page_claims: dict[int, set[str]] = {}
        for page, entry in enumerate(slide_copy, 1):
            claim_ids = entry.get("claim_ids")
            if not isinstance(claim_ids, list):
                add(result, "errors", f"slide-copy ledger page {page} claim_ids must be an explicit string array")
                ledger_page_claims[page] = set()
                continue
            ledger_page_claims[page] = {value for value in claim_ids if isinstance(value, str)}
            if not ledger_page_claims[page] and not normalize_text(entry.get("claim_not_applicable_reason")):
                add(
                    result,
                    "errors",
                    f"slide-copy ledger page {page} requires claim_ids or claim_not_applicable_reason",
                )
            for claim_id in ledger_page_claims[page]:
                claim = claims_by_id.get(claim_id)
                if claim is None:
                    continue
                used_on_slides = claim.get("used_on_slides")
                if not isinstance(used_on_slides, list) or page not in used_on_slides:
                    add(
                        result,
                        "errors",
                        f"claim {claim_id} does not reciprocally record slide {page} in used_on_slides",
                    )
                slide_evidence = evidence_for_page(claim, "slide_evidence", page)
                if not slide_evidence:
                    statement = normalize_text(claim.get("statement"))
                    slide_evidence = [statement] if statement else []
                slide_text = "\n".join(entry_text(entry, SLIDE_TEXT_FIELDS))
                missing_evidence = [
                    snippet for snippet in slide_evidence
                    if not text_contains(slide_text, snippet)
                ]
                if not slide_evidence or missing_evidence:
                    add(
                        result,
                        "errors",
                        f"slide-copy ledger page {page} does not contain claim {claim_id} evidence: {missing_evidence or '[missing]'}",
                    )
        for claim in claims:
            claim_id = claim.get("id")
            used_on_slides = claim.get("used_on_slides")
            if not isinstance(claim_id, str) or not isinstance(used_on_slides, list):
                continue
            for page in used_on_slides:
                if not isinstance(page, int) or isinstance(page, bool):
                    continue
                if page not in ledger_page_claims:
                    add(result, "errors", f"claim {claim_id} references slide {page} outside the slide-copy ledger")
                elif claim_id not in ledger_page_claims[page]:
                    add(
                        result,
                        "errors",
                        f"claim {claim_id} records slide {page}, but that slide omits the claim_id",
                    )
        if notes_ledger is not None:
            if len(notes_ledger) != len(slide_copy):
                add(result, "errors", "speaker-notes and slide-copy claim mapping page counts differ")
            for page, note in enumerate(notes_ledger, 1):
                note_ids = note.get("claim_ids")
                if not isinstance(note_ids, list):
                    add(result, "errors", f"speaker-notes ledger page {page} claim_ids must be an explicit string array")
                    continue
                note_claims = {value for value in note_ids if isinstance(value, str)}
                extra = sorted(note_claims - ledger_page_claims.get(page, set()))
                if extra:
                    add(result, "errors", f"speaker-notes ledger page {page} adds slide-unapproved claim IDs: {extra}")
                if not note_claims and not normalize_text(note.get("claim_not_applicable_reason")):
                    add(
                        result,
                        "errors",
                        f"speaker-notes ledger page {page} requires claim_ids or claim_not_applicable_reason",
                    )
                note_text = "\n".join(entry_text(note, NOTE_TEXT_FIELDS))
                for claim_id in note_claims:
                    claim = claims_by_id.get(claim_id)
                    if claim is None:
                        continue
                    note_evidence = evidence_for_page(claim, "notes_evidence", page)
                    if not note_evidence:
                        note_evidence = evidence_for_page(claim, "slide_evidence", page)
                    if not note_evidence:
                        statement = normalize_text(claim.get("statement"))
                        note_evidence = [statement] if statement else []
                    missing_note_evidence = [
                        snippet for snippet in note_evidence
                        if not text_contains(note_text, snippet)
                    ]
                    if not note_evidence or missing_note_evidence:
                        add(
                            result,
                            "errors",
                            f"speaker-notes ledger page {page} does not contain claim {claim_id} evidence: {missing_note_evidence or '[missing]'}",
                        )


def validate_report_claim_coverage(
    report_info: DocxInfo | None,
    claim_data: dict[str, Any] | None,
    result: dict[str, Any],
) -> None:
    if report_info is None or claim_data is None:
        return
    claims = claim_data.get("claims") if isinstance(claim_data.get("claims"), list) else []
    report_claims: list[dict[str, Any]] = []
    for index, claim in enumerate(claims, 1):
        if not isinstance(claim, dict):
            continue
        used = claim.get("used_in_report")
        if not isinstance(used, bool):
            add(result, "errors", f"claim-ledger claim {index} used_in_report must be Boolean for report delivery")
            continue
        if used:
            report_claims.append(claim)
    if not report_claims:
        add(result, "errors", "report delivery requires at least one resolved claim with used_in_report=true")
        return
    for claim in report_claims:
        claim_id = claim.get("id", "(unknown)")
        evidence = claim.get("report_evidence")
        snippets = flatten_text_values(evidence) if evidence is not None else []
        if not snippets:
            statement = normalize_text(claim.get("statement"))
            snippets = [statement] if statement else []
        if not snippets:
            add(result, "errors", f"claim {claim_id} has no report_evidence or statement to reconcile")
            continue
        missing = [snippet for snippet in snippets if not text_contains(report_info.text, snippet)]
        if missing:
            add(result, "errors", f"report does not contain claim {claim_id} evidence: {missing}")
    result["metrics"]["report_claims_reconciled"] = len(report_claims)


def validate_project_state(
    path: Path,
    result: dict[str, Any],
    source_manifest: dict[str, Any] | None,
    style_options_path: Path | None,
) -> dict[str, Any] | None:
    data = parse_json_file(path.expanduser().resolve(), "project state", result)
    if not isinstance(data, dict):
        return None
    state_file = path.expanduser().resolve()
    if project_state_manifest_errors is None or project_state_detect_drift is None:
        add(result, "errors", "project state contract module could not be loaded")
        return data
    contract_errors = project_state_manifest_errors(data)
    for message in contract_errors:
        add(result, "errors", f"project state contract: {message}")
    project_dir_value = data.get("project_dir")
    if (
        not isinstance(project_dir_value, str)
        or Path(project_dir_value).expanduser().resolve() != state_file.parent
    ):
        add(result, "errors", "project state project_dir does not match the state file location")
    if not contract_errors:
        try:
            drift = project_state_detect_drift(copy.deepcopy(data))
        except (OSError, RuntimeError, KeyError, TypeError, ValueError) as exc:
            add(result, "errors", f"project state read-only drift verification failed: {exc}")
        else:
            if drift:
                add(result, "errors", f"project state contains current artifact drift: {drift}")
    if data.get("schema_version") != 2:
        add(result, "errors", "project state requires schema_version: 2")
    policies = data.get("policies")
    if not isinstance(policies, dict):
        add(result, "errors", "project state requires structured policies")
        policies = {}
    if policies.get("source_policy") not in {"faithful", "verify-update"}:
        add(result, "errors", "project state policies.source_policy is invalid")
    if policies.get("data_handling") not in {"standard", "local-only"}:
        add(result, "errors", "project state policies.data_handling is invalid")
    if policies.get("ocr_policy") != "offline" or policies.get("ask_for_ocr_token") is not False:
        add(result, "errors", "project state must record offline OCR with ask_for_ocr_token=false")
    if source_manifest is not None:
        for key in ("source_policy", "data_handling"):
            if policies.get(key) != source_manifest.get(key):
                add(result, "errors", f"project state {key} does not match source-manifest")
    if style_options_path is not None:
        style_data = parse_json_file(style_options_path.expanduser().resolve(), "style-options state linkage", result)
        style_stage = (data.get("stages") or {}).get("style-options") if isinstance(data.get("stages"), dict) else None
        state_selection = style_stage.get("selection") if isinstance(style_stage, dict) else None
        if isinstance(style_data, dict) and state_selection != style_data.get("selected_id"):
            add(result, "errors", "project state selected style does not match style-options selected_id")
    history = data.get("history")
    verified = [entry for entry in history if isinstance(entry, dict) and entry.get("event") == "verified"] if isinstance(history, list) else []
    if not verified:
        add(result, "errors", "project state has no recorded verify checkpoint")
    else:
        last_verify = verified[-1]
        details = verified[-1].get("details") if isinstance(verified[-1].get("details"), dict) else {}
        if not isinstance(last_verify.get("id"), int) or isinstance(last_verify.get("id"), bool):
            add(result, "errors", "project state latest verify checkpoint has an invalid history id")
        if details.get("drift_count") != 0:
            add(result, "errors", "project state latest verify checkpoint contains drift")
        state_revision = data.get("state_revision")
        verify_revision = last_verify.get("revision")
        if not isinstance(state_revision, int) or not isinstance(verify_revision, int) or verify_revision != state_revision:
            add(result, "errors", "project state verify checkpoint is stale; run verify immediately before final validation")
    stages = data.get("stages") if isinstance(data.get("stages"), dict) else {}
    seen_accepted_qa: set[Path] = set()
    for stage_name in data.get("targets", []) if isinstance(data.get("targets"), list) else []:
        stage = stages.get(stage_name) if isinstance(stages, dict) else None
        if not isinstance(stage, dict):
            continue
        for warning in stage.get("warnings", []) if isinstance(stage.get("warnings"), list) else []:
            if normalize_text(warning):
                add(result, "warnings", f"project-state {stage_name}: {normalize_text(warning)}")
        pages = stage.get("pages") if isinstance(stage.get("pages"), dict) else {}
        for page_id, page in pages.items():
            if not isinstance(page, dict):
                continue
            for warning in page.get("warnings", []) if isinstance(page.get("warnings"), list) else []:
                if normalize_text(warning):
                    add(
                        result,
                        "warnings",
                        f"project-state {stage_name} page {page_id}: {normalize_text(warning)}",
                    )
        # A completed phase's accepted QA is historical authority.  Carry its
        # warnings into every later final validation so they cannot disappear
        # merely because the stage's free-form warnings list was left empty.
        if stage.get("status") != "completed":
            continue
        stage_artifacts = stage.get("artifacts") if isinstance(stage.get("artifacts"), dict) else {}
        for key, record in stage_artifacts.items():
            if not isinstance(record, dict) or not isinstance(record.get("path"), str):
                continue
            qa_path = Path(record["path"]).expanduser().resolve()
            if (
                qa_path in seen_accepted_qa
                or ("qa" not in str(key).casefold() and qa_path.name != "qa-report.json")
                or not qa_path.is_file()
            ):
                continue
            seen_accepted_qa.add(qa_path)
            try:
                qa_payload = json.loads(qa_path.read_text(encoding="utf-8"))
            except (OSError, UnicodeError, json.JSONDecodeError):
                continue
            if not isinstance(qa_payload, dict) or qa_payload.get("tool") != "ppt-gen.validate_delivery":
                continue
            if qa_payload.get("validation_contract_version", 1) == 1 and any(
                isinstance(artifact, dict) and artifact.get("role") == "editable-pptx"
                for artifact in qa_payload.get("artifacts", [])
            ):
                add(result, "warnings", f"legacy editable QA in {stage_name} predates guarded offline execution evidence; new/revalidated editable outputs require the current gate")
            for warning in qa_payload.get("warnings", []) if isinstance(qa_payload.get("warnings"), list) else []:
                if normalize_text(warning):
                    add(
                        result,
                        "warnings",
                        f"accepted qa-report {stage_name}: {normalize_text(warning)}",
                    )
    return data


def state_known_records(data: dict[str, Any] | None) -> list[dict[str, Any]]:
    if not isinstance(data, dict):
        return []
    records: list[dict[str, Any]] = []
    for container_key in ("inputs", "deliverable_artifacts", "artifacts"):
        container = data.get(container_key)
        if isinstance(container, dict):
            records.extend(record for record in container.values() if isinstance(record, dict))
    stages = data.get("stages")
    if isinstance(stages, dict):
        for stage in stages.values():
            if not isinstance(stage, dict):
                continue
            artifacts = stage.get("artifacts")
            if isinstance(artifacts, dict):
                records.extend(record for record in artifacts.values() if isinstance(record, dict))
            pages = stage.get("pages")
            if isinstance(pages, dict):
                for page in pages.values():
                    page_artifacts = page.get("artifacts") if isinstance(page, dict) else None
                    if isinstance(page_artifacts, dict):
                        records.extend(record for record in page_artifacts.values() if isinstance(record, dict))
    return records


def validate_state_artifact_linkage(
    state_data: dict[str, Any] | None,
    artifacts: list[dict[str, Any]],
    result: dict[str, Any],
) -> None:
    if state_data is None:
        return
    known = state_known_records(state_data)
    requested = {
        role for role in state_data.get("deliverables", [])
        if isinstance(role, str)
    }
    deliverable_records = (
        state_data.get("deliverable_artifacts")
        if isinstance(state_data.get("deliverable_artifacts"), dict)
        else {}
    )
    stages = state_data.get("stages") if isinstance(state_data.get("stages"), dict) else {}
    for artifact in artifacts:
        if artifact.get("role") in {"project-state"}:
            continue
        path = artifact.get("path")
        digest = artifact.get("sha256")
        if not path or not digest:
            continue
        delivered_contract = CLI_DELIVERABLE_ROLES.get(str(artifact.get("role")))
        if delivered_contract is not None:
            deliverable_role, expected_stage = delivered_contract
            if deliverable_role in requested:
                declared = deliverable_records.get(deliverable_role)
                declared_matches = (
                    isinstance(declared, dict)
                    and declared.get("stage") == expected_stage
                    and declared.get("status", "current") == "current"
                    and isinstance(declared.get("path"), str)
                    and Path(declared["path"]).expanduser().resolve()
                    == Path(path).expanduser().resolve()
                    and isinstance(declared.get("sha256"), str)
                    and declared["sha256"].casefold() == str(digest).casefold()
                )
                if not declared_matches:
                    add(
                        result,
                        "errors",
                        f"project state requested deliverable {deliverable_role!r} is not registered at its current path/hash",
                    )
                    continue
            else:
                if matching_import(state_data, artifact):
                    continue
                stage_item = stages.get(expected_stage) if isinstance(stages, dict) else None
                stage_records = stage_item.get("artifacts") if isinstance(stage_item, dict) else None
                targets = state_data.get("targets") if isinstance(state_data.get("targets"), list) else []
                stage_is_active_authority = (
                    expected_stage in targets
                    and isinstance(stage_item, dict)
                    and stage_item.get("status") not in {"skipped", "cancelled", "failed"}
                )
                internal_match = any(
                    isinstance(record, dict)
                    and record.get("status", "current") == "current"
                    and isinstance(record.get("path"), str)
                    and Path(record["path"]).expanduser().resolve()
                    == Path(path).expanduser().resolve()
                    and isinstance(record.get("sha256"), str)
                    and record["sha256"].casefold() == str(digest).casefold()
                    for record in (stage_records or {}).values()
                )
                if not stage_is_active_authority or not internal_match:
                    add(
                        result,
                        "errors",
                        f"internal artifact {artifact['role']} is not registered in expected state stage {expected_stage}",
                    )
                    continue
        linked = any(
            Path(str(record.get("path", ""))).expanduser().resolve() == Path(path).expanduser().resolve()
            and isinstance(record.get("sha256"), str)
            and record.get("sha256", "").casefold() == str(digest).casefold()
            and record.get("status", "current") == "current"
            for record in known
        )
        if not linked:
            add(result, "errors", f"project state does not register current artifact {artifact['role']}: {path}")


def validate_editable_handoff(
    path: Path,
    result: dict[str, Any],
    source_manifest_path: Path | None,
    slide_copy_path: Path | None,
    slide_manifest_path: Path | None,
    source_images: list[ImageRecord] | None,
) -> dict[str, Any] | None:
    resolved = path.expanduser().resolve()
    data = parse_json_file(resolved, "editable-handoff", result)
    if not isinstance(data, dict):
        return None
    if data.get("schema_version") != 1:
        add(result, "errors", "editable-handoff requires schema_version: 1")
    if data.get("ocr_policy") != "offline":
        add(result, "errors", "editable-handoff ocr_policy must be 'offline'")
    if data.get("ask_for_ocr_token") is not False:
        add(result, "errors", "editable-handoff ask_for_ocr_token must be false")
    data_handling = data.get("data_handling")
    if data_handling not in {"standard", "local-only"}:
        add(result, "errors", "editable-handoff data_handling must be 'standard' or 'local-only'")
    if data.get("source_policy") not in {"faithful", "verify-update"}:
        add(result, "errors", "editable-handoff source_policy must be 'faithful' or 'verify-update'")
    image_backend = data.get("image_backend")
    if image_backend not in {"builtin-allowed", "disabled"}:
        add(result, "errors", "editable-handoff image_backend must be 'builtin-allowed' or 'disabled'")
    if data_handling == "local-only" and image_backend != "disabled":
        add(result, "errors", "editable-handoff local-only handling requires image_backend='disabled'")
    if data.get("implementation") != "image-to-editable-ppt":
        add(result, "errors", "editable-handoff implementation must be 'image-to-editable-ppt'")

    authoritative_paths: dict[str, Path] = {}
    for key in ("source_manifest", "slide_copy_ledger", "slide_manifest"):
        value = data.get(key)
        if not isinstance(value, str) or not Path(value).expanduser().is_absolute():
            add(result, "errors", f"editable-handoff {key} must be an absolute path")
            continue
        target = Path(value).expanduser().resolve()
        authoritative_paths[key] = target
        if not target.is_file():
            add(result, "errors", f"editable-handoff {key} is missing: {target}")
    source_manifest = authoritative_paths.get("source_manifest")
    if source_manifest is not None and source_manifest.is_file():
        source_data = validate_source_manifest(source_manifest, result)
        if isinstance(source_data, dict):
            if source_data.get("data_handling") != data_handling:
                add(result, "errors", "editable-handoff data_handling does not match source-manifest")
            if source_data.get("source_policy") != data.get("source_policy"):
                add(result, "errors", "editable-handoff source_policy does not match source-manifest")
    expected_source = source_manifest_path.expanduser().resolve() if source_manifest_path else None
    if expected_source is not None and authoritative_paths.get("source_manifest") != expected_source:
        add(result, "errors", "editable-handoff source_manifest does not match --source-manifest")
    expected_ledger = slide_copy_path.expanduser().resolve() if slide_copy_path else None
    expected_manifest = slide_manifest_path.expanduser().resolve() if slide_manifest_path else None
    if expected_ledger is not None and authoritative_paths.get("slide_copy_ledger") != expected_ledger:
        add(result, "errors", "editable-handoff slide_copy_ledger does not match --slide-copy-ledger")
    if expected_manifest is not None and authoritative_paths.get("slide_manifest") != expected_manifest:
        add(result, "errors", "editable-handoff slide_manifest does not match --slide-manifest")

    source_entries = data.get("source_images")
    handoff_paths: list[Path] = []
    if not isinstance(source_entries, list) or not source_entries:
        add(result, "errors", "editable-handoff source_images must be a nonempty ordered array")
    else:
        for index, entry in enumerate(source_entries, 1):
            path_value: Any = entry
            expected_hash: Any = None
            if isinstance(entry, dict):
                path_value = entry.get("path") or entry.get("image")
                expected_hash = entry.get("sha256")
                if not isinstance(expected_hash, str) or not expected_hash:
                    add(result, "errors", f"editable-handoff source_images[{index}] requires sha256")
            if not isinstance(path_value, str) or not Path(path_value).expanduser().is_absolute():
                add(result, "errors", f"editable-handoff source_images[{index}] path must be absolute")
                continue
            image_path = Path(path_value).expanduser().resolve()
            handoff_paths.append(image_path)
            if not image_path.is_file():
                add(result, "errors", f"editable-handoff source image {index} is missing: {image_path}")
                continue
            try:
                decode_image_path(image_path)
            except Exception as exc:
                add(result, "errors", f"editable-handoff source image {index} cannot be decoded: {exc}")
            if not isinstance(expected_hash, str) or expected_hash.casefold() != sha256_file(image_path).casefold():
                add(result, "errors", f"editable-handoff source image {index} SHA-256 mismatch")
    if source_images is not None:
        expected_paths = [record.path.resolve() for record in source_images]
        if handoff_paths != expected_paths:
            add(result, "errors", "editable-handoff source_images do not match the validated slide-manifest order")

    handoff_hashes = data.get("hashes") if isinstance(data.get("hashes"), dict) else {}
    bindings = {
        "source_manifest": authoritative_paths.get("source_manifest"),
        "slide_copy_ledger": authoritative_paths.get("slide_copy_ledger"),
        "slide_manifest": authoritative_paths.get("slide_manifest"),
    }
    for key, target in bindings.items():
        expected_hash = data.get(f"{key}_sha256") or handoff_hashes.get(key)
        if expected_hash is None:
            add(result, "errors", f"editable-handoff {key} requires a stored SHA-256 binding")
        elif target is not None and target.is_file() and (
            not isinstance(expected_hash, str) or expected_hash.casefold() != sha256_file(target).casefold()
        ):
            add(result, "errors", f"editable-handoff {key} SHA-256 mismatch")

    prompt_tail = normalize_text(data.get("worker_prompt_tail"))
    prompt_key = prompt_tail.casefold()
    offline_ok = "offline" in prompt_key or "离线" in prompt_key
    token_terms = bool(re.search(r"(?:never|do not|don't|no|不|不要|禁)[^。.;]{0,40}(?:paddleocr|ocr)[^。.;]{0,20}token", prompt_key))
    if not offline_ok or not token_terms:
        add(result, "errors", "editable-handoff worker_prompt_tail must explicitly require offline OCR and forbid requesting an OCR token")
    return data


def embedded_notes_texts(
    path: Path,
    result: dict[str, Any],
    label: str,
) -> list[str] | None:
    package = inspect_opc_package(path, result, label, "ppt/presentation.xml")
    if package is None:
        return None
    parts = ordered_slide_parts(package, result, label)
    output: list[str] = []
    for page, slide_part in enumerate(parts, 1):
        note_relationships = [
            rel
            for rel in package.rels.get(slide_part, {}).values()
            if rel.rel_type.rstrip("/").endswith("/notesSlide") and rel.resolved
        ]
        if len(note_relationships) > 1:
            add(result, "errors", f"{label} slide {page} has multiple notesSlide relationships")
        if not note_relationships:
            output.append("")
            continue
        note_root = package.roots.get(note_relationships[0].resolved or "")
        if note_root is None:
            add(result, "errors", f"{label} slide {page} notes XML is missing or malformed")
            output.append("")
            continue
        output.append("\n".join(ppt_paragraphs(note_root)))
    return output


def validate_embedded_notes_pptx(
    path: Path,
    info: PptxInfo | None,
    notes_ledger: list[dict[str, Any]] | None,
    title_entries: list[dict[str, Any]] | None,
    page_manifest: list[dict[str, Any]] | None,
    result: dict[str, Any],
) -> None:
    if info is None:
        return
    note_texts = embedded_notes_texts(path, result, "ppt_notes_pptx")
    if note_texts is None:
        return
    if notes_ledger is None:
        add(result, "errors", "--ppt-notes-pptx requires --speaker-notes-ledger")
        return
    if len(note_texts) != len(notes_ledger):
        add(result, "errors", f"embedded notes pages {len(note_texts)} != speaker-notes ledger pages {len(notes_ledger)}")
        return
    if title_entries is None:
        add(result, "errors", "--ppt-notes-pptx requires --slide-copy-ledger for exact deck-title mapping")
    for page, (slide, note_text, ledger_entry) in enumerate(zip(info.slides, note_texts, notes_ledger), 1):
        if not normalize_text(note_text):
            add(result, "errors", f"ppt_notes_pptx slide {page} has no embedded speaker notes")
            continue
        expected_title = normalize_text(
            (title_entries[page - 1] if title_entries and page <= len(title_entries) else ledger_entry).get("title", "")
        )
        if expected_title and slide.title:
            if compact_text(slide.title) != compact_text(expected_title):
                add(result, "errors", f"ppt_notes_pptx slide {page} title mismatch: {slide.title!r} != {expected_title!r}")
        elif expected_title:
            manifest_title = normalize_text(
                page_manifest[page - 1].get("title", "")
                if page_manifest and page <= len(page_manifest)
                else ""
            )
            if compact_text(manifest_title) != compact_text(expected_title):
                add(result, "errors", f"ppt_notes_pptx image-only slide {page} requires a matching --slide-manifest title")
        ledger_title = normalize_text(ledger_entry.get("title", ""))
        if expected_title and ledger_title and compact_text(ledger_title) != compact_text(expected_title):
            add(result, "errors", f"speaker-notes ledger title mismatch on embedded-notes page {page}")
        missing = [value for value in entry_text(ledger_entry, NOTE_TEXT_FIELDS) if not text_contains(note_text, value)]
        if missing:
            add(result, "errors", f"ppt_notes_pptx slide {page} embedded notes are missing ledger text: {missing}")
    result["metrics"]["ppt_notes_pages"] = len(note_texts)


def required_absolute_file(
    value: Any,
    label: str,
    result: dict[str, Any],
) -> Path | None:
    if not isinstance(value, str) or not Path(value).expanduser().is_absolute():
        add(result, "errors", f"{label} must be an absolute path")
        return None
    path = Path(value).expanduser().resolve()
    if not path.is_file():
        add(result, "errors", f"{label} is missing: {path}")
        return None
    if path.stat().st_size == 0:
        add(result, "errors", f"{label} is empty: {path}")
        return None
    return path


def validate_revision_checkpoints(
    path: Path,
    original_path: Path,
    revised_path: Path,
    original_info: PptxInfo | None,
    revised_info: PptxInfo | None,
    claim_data: dict[str, Any] | None,
    result: dict[str, Any],
) -> None:
    resolved = path.expanduser().resolve()
    data = parse_json_file(resolved, "revision-checkpoints", result)
    if not isinstance(data, dict):
        return
    if data.get("schema_version") != 1:
        add(result, "errors", "revision-checkpoints requires schema_version: 1")
    if original_path == revised_path:
        add(result, "errors", "deck revision must preserve the original at a different path")
    original_hash = sha256_file(original_path) if original_path.is_file() else ""
    revised_hash = sha256_file(revised_path) if revised_path.is_file() else ""
    if original_hash and revised_hash and original_hash == revised_hash:
        add(result, "errors", "revised PPTX is byte-identical to the original")
    revision_hashes = data.get("hashes") if isinstance(data.get("hashes"), dict) else {}
    for key, object_key, expected_path, expected_hash in (
        ("original_pptx", "original", original_path, original_hash),
        ("revised_pptx", "revised", revised_path, revised_hash),
    ):
        authority = data.get(object_key) if isinstance(data.get(object_key), dict) else {}
        value = data.get(key) or authority.get("path")
        if not isinstance(value, str) or Path(value).expanduser().resolve() != expected_path:
            add(result, "errors", f"revision-checkpoints {key} does not match the command-line artifact")
        stored_hash = data.get(f"{key}_sha256") or revision_hashes.get(key) or authority.get("sha256")
        if not isinstance(stored_hash, str) or stored_hash.casefold() != expected_hash.casefold():
            add(result, "errors", f"revision-checkpoints {key} SHA-256 is missing or mismatched")

    if original_info is None or revised_info is None:
        return
    original_count = len(original_info.slides)
    revised_count = len(revised_info.slides)
    if (original_info.width, original_info.height) != (revised_info.width, revised_info.height):
        add(result, "errors", "deck revision changed slide geometry")
    if original_count != revised_count and data.get("allow_slide_count_change") is not True:
        add(result, "errors", f"deck revision changed slide count {original_count}->{revised_count} without explicit approval")

    pages = data.get("pages") or data.get("checkpoints")
    if not isinstance(pages, list) or not pages:
        add(result, "errors", "revision-checkpoints requires a nonempty pages array")
        pages = []
    changed_pages: set[int] = set()
    revision_evidence_paths: set[Path] = set()
    checkpoint_evidence_by_page: dict[int, list[Path]] = {}
    known_claims = {
        claim.get("id"): claim
        for claim in (claim_data or {}).get("claims", [])
        if isinstance(claim, dict) and isinstance(claim.get("id"), str)
    }
    seen_checkpoint_ids: set[str] = set()
    seen_object_keys: set[tuple[int, str]] = set()
    for index, checkpoint in enumerate(pages, 1):
        if not isinstance(checkpoint, dict):
            add(result, "errors", f"revision checkpoint {index} is not an object")
            continue
        page = entry_page_number(checkpoint)
        if page is None or page < 1 or page > revised_count:
            add(result, "errors", f"revision checkpoint {index} has invalid page {page!r}")
            continue
        changed_pages.add(page)
        checkpoint_id = checkpoint.get("id")
        object_id = checkpoint.get("object_id")
        if not isinstance(checkpoint_id, str) or not checkpoint_id:
            add(result, "errors", f"revision checkpoint {index} requires id")
        elif checkpoint_id in seen_checkpoint_ids:
            add(result, "errors", f"revision-checkpoints repeats id {checkpoint_id!r}")
        else:
            seen_checkpoint_ids.add(checkpoint_id)
        if not isinstance(object_id, str) or not object_id:
            add(result, "errors", f"revision checkpoint {index} requires object_id")
        elif (page, object_id) in seen_object_keys:
            add(result, "errors", f"revision-checkpoints repeats page/object ({page}, {object_id!r})")
        else:
            seen_object_keys.add((page, object_id))
        requested = checkpoint.get("requested_changes")
        if requested is None and normalize_text(checkpoint.get("request")):
            requested = [checkpoint.get("request")]
        if not isinstance(requested, list) or not requested or not all(normalize_text(item) for item in requested):
            add(result, "errors", f"revision checkpoint page {page} requires requested_changes")
        if checkpoint.get("status") != "passed":
            add(result, "errors", f"revision checkpoint page {page} status must be 'passed'")
        checkpoint_claim_ids = checkpoint.get("claim_ids")
        if not isinstance(checkpoint_claim_ids, list) or not all(
            isinstance(value, str) for value in checkpoint_claim_ids
        ):
            add(result, "errors", f"revision checkpoint page {page} claim_ids must be an explicit string array")
            checkpoint_claim_ids = []
        if not checkpoint_claim_ids and not normalize_text(checkpoint.get("claim_not_applicable_reason")):
            add(
                result,
                "errors",
                f"revision checkpoint page {page} requires claim_ids or claim_not_applicable_reason",
            )
        unknown_claims = sorted(set(checkpoint_claim_ids) - set(known_claims))
        if unknown_claims:
            add(result, "errors", f"revision checkpoint page {page} references unknown claim IDs: {unknown_claims}")
        revised_page_text = revised_info.slides[page - 1].visible_text
        for claim_id in checkpoint_claim_ids:
            claim = known_claims.get(claim_id)
            if claim is None:
                continue
            used_on_slides = claim.get("used_on_slides")
            if not isinstance(used_on_slides, list) or page not in used_on_slides:
                add(result, "errors", f"revision claim {claim_id} does not record page {page} in used_on_slides")
            evidence = claim.get("slide_evidence")
            page_evidence = None
            if isinstance(evidence, dict):
                page_evidence = evidence.get(str(page), evidence.get(page))
            snippets = flatten_text_values(page_evidence)
            if not snippets:
                statement = normalize_text(claim.get("statement"))
                snippets = [statement] if statement else []
            missing_claim_text = [snippet for snippet in snippets if not text_contains(revised_page_text, snippet)]
            if not snippets or missing_claim_text:
                add(
                    result,
                    "errors",
                    f"revision checkpoint page {page} does not contain claim {claim_id} evidence: {missing_claim_text or '[missing]'}",
                )
        change_type = normalize_text(checkpoint.get("change_type") or "text").casefold()
        if change_type not in {"text", "visual", "geometry", "media"}:
            add(
                result,
                "errors",
                f"revision checkpoint page {page} change_type must be text, visual, geometry, or media",
            )
            change_type = "text"
        expected_after = checkpoint.get("expected_after") or checkpoint.get("new_value")
        if expected_after is None:
            add(result, "errors", f"revision checkpoint page {page} requires expected_after/new_value for machine reconciliation")
        elif change_type == "text":
            original_slide = original_info.slides[page - 1]
            revised_slide = revised_info.slides[page - 1]
            normalized_object_id = normalize_text(object_id)
            object_key = normalized_object_id
            if normalized_object_id.casefold() in {"title", "title-01", "slide-title"}:
                object_key = "title"
            revised_object_text = revised_slide.text_objects.get(object_key)
            original_object_text = original_slide.text_objects.get(object_key)
            if revised_object_text is None or original_object_text is None:
                known_objects = sorted(
                    key for key in set(original_slide.text_objects) | set(revised_slide.text_objects)
                    if not key.startswith("id:") and not key.startswith("name:")
                )
                add(
                    result,
                    "errors",
                    f"revision checkpoint page {page} object_id {object_id!r} does not bind a stable native text object; known={known_objects}",
                )
                revised_object_text = ""
            expected_values = flatten_text_values(expected_after)
            if not expected_values or any(not text_contains(revised_object_text, value) for value in expected_values):
                add(result, "errors", f"revision checkpoint page {page} expected_after is absent from the bound revised text object")
            expected_before = checkpoint.get("expected_before") or checkpoint.get("old_value")
            if expected_before is None:
                add(result, "errors", f"revision checkpoint page {page} text change requires expected_before/old_value")
            else:
                before_values = flatten_text_values(expected_before)
                if not before_values or any(not text_contains(original_object_text or "", value) for value in before_values):
                    add(result, "errors", f"revision checkpoint page {page} expected_before is absent from the bound original text object")
                if any(text_contains(revised_object_text, value) for value in before_values):
                    add(result, "errors", f"revision checkpoint page {page} old value remains in the bound revised text object")
            if compact_text(original_object_text or "") == compact_text(revised_object_text):
                add(result, "errors", f"revision checkpoint page {page} bound text object did not change")
            if object_key == "title" and compact_text(revised_slide.title) != compact_text(revised_object_text):
                add(result, "errors", f"revision checkpoint page {page} title object does not match the actual slide title")
        else:
            if not flatten_text_values(expected_after):
                add(result, "errors", f"revision checkpoint page {page} requires a visual expected_after description")
            original_slide = original_info.slides[page - 1]
            revised_slide = revised_info.slides[page - 1]
            before_fingerprint = checkpoint.get("before_slide_fingerprint") or checkpoint.get("before_fingerprint")
            after_fingerprint = checkpoint.get("after_slide_fingerprint") or checkpoint.get("after_fingerprint")
            if (
                not isinstance(before_fingerprint, str)
                or before_fingerprint.casefold() != original_slide.semantic_fingerprint.casefold()
            ):
                add(result, "errors", f"revision checkpoint page {page} before slide fingerprint is missing or mismatched")
            if (
                not isinstance(after_fingerprint, str)
                or after_fingerprint.casefold() != revised_slide.semantic_fingerprint.casefold()
            ):
                add(result, "errors", f"revision checkpoint page {page} after slide fingerprint is missing or mismatched")
            if original_slide.semantic_fingerprint == revised_slide.semantic_fingerprint:
                add(result, "errors", f"revision checkpoint page {page} records a {change_type} change but slide semantics did not change")
            if change_type == "media":
                original_media = sorted(original_slide.picture_fingerprints.values())
                revised_media = sorted(revised_slide.picture_fingerprints.values())
                if original_media == revised_media:
                    add(result, "errors", f"revision checkpoint page {page} records a media change but picture fingerprints did not change")
        evidence = checkpoint.get("evidence")
        if not isinstance(evidence, list) or not evidence:
            add(result, "errors", f"revision checkpoint page {page} requires evidence files")
        else:
            for evidence_index, item in enumerate(evidence, 1):
                evidence_path = validate_bound_evidence_file(
                    item,
                    resolved.parent,
                    f"revision page {page} evidence {evidence_index}",
                    result,
                    allowed_suffixes=IMAGE_EXTENSIONS,
                )
                if evidence_path is not None:
                    if evidence_path in revision_evidence_paths:
                        add(result, "errors", f"revision evidence file is reused across checkpoints: {evidence_path}")
                    revision_evidence_paths.add(evidence_path)
                    checkpoint_evidence_by_page.setdefault(page, []).append(evidence_path)
        critical = checkpoint.get("unchanged_critical_fields", [])
        if not isinstance(critical, list):
            add(result, "errors", f"revision checkpoint page {page} unchanged_critical_fields must be an array")
            critical = []
        if page <= original_count:
            original_text = original_info.slides[page - 1].visible_text
            revised_text = revised_info.slides[page - 1].visible_text
            for field_value in critical:
                if not text_contains_exact_field(original_text, str(field_value)) or not text_contains_exact_field(
                    revised_text, str(field_value)
                ):
                    add(result, "errors", f"revision page {page} did not preserve critical field {field_value!r}")

    unchanged_checks = data.get("unchanged_checks", [])
    if not isinstance(unchanged_checks, list):
        add(result, "errors", "revision-checkpoints unchanged_checks must be an array")
        unchanged_checks = []
    for index, check in enumerate(unchanged_checks, 1):
        if not isinstance(check, dict):
            add(result, "errors", f"revision unchanged_check {index} is not an object")
            continue
        page = entry_page_number(check)
        fields = check.get("fields")
        if page is None or not 1 <= page <= min(original_count, revised_count):
            add(result, "errors", f"revision unchanged_check {index} has invalid page {page!r}")
            continue
        if not isinstance(fields, list) or not fields:
            add(result, "errors", f"revision unchanged_check page {page} requires fields")
            fields = []
        if check.get("status") != "passed":
            add(result, "errors", f"revision unchanged_check page {page} status must be 'passed'")
        evidence = check.get("evidence")
        if not isinstance(evidence, list) or not evidence:
            add(result, "errors", f"revision unchanged_check page {page} requires evidence")
        else:
            for evidence_index, item in enumerate(evidence, 1):
                validate_bound_evidence_file(
                    item,
                    resolved.parent,
                    f"revision unchanged_check page {page} evidence {evidence_index}",
                    result,
                )
        for field_value in fields:
            if not text_contains_exact_field(original_info.slides[page - 1].visible_text, str(field_value)) or not text_contains_exact_field(
                revised_info.slides[page - 1].visible_text, str(field_value)
            ):
                add(result, "errors", f"revision unchanged_check page {page} did not preserve {field_value!r}")
    if original_count > len(changed_pages) and not unchanged_checks:
        add(result, "errors", "revision-checkpoints requires unchanged_checks for pages outside the requested edit scope")

    contact_sheet_path = validate_bound_evidence_file(
        data.get("contact_sheet"),
        resolved.parent,
        "revision-checkpoints contact_sheet",
        result,
        allowed_suffixes=IMAGE_EXTENSIONS,
    )
    if contact_sheet_path is not None and contact_sheet_path in revision_evidence_paths:
        add(result, "errors", "revision contact sheet must be distinct from per-object/page evidence")

    original_render_pages = independently_render_artifact_pages(original_path, result, "deck-revision original")
    revised_render_pages = independently_render_artifact_pages(revised_path, result, "deck-revision revised")
    if original_render_pages is None or revised_render_pages is None:
        add(result, "errors", "deck-revision could not independently render original and revised decks")
    else:
        for page in sorted(changed_pages):
            required = [original_render_pages[page - 1], revised_render_pages[page - 1]]
            page_evidence = checkpoint_evidence_by_page.get(page, [])
            if not page_evidence:
                continue
            matched_pair = False
            for candidate in page_evidence:
                before_match = False
                after_match = False
                try:
                    _, _, candidate_pixels, candidate_phash = decode_image_path(candidate)
                    for bound_index, bound in enumerate(required):
                        _, _, bound_pixels, bound_phash = decode_image_path(bound)
                        matched = candidate_pixels == bound_pixels or hamming_distance(candidate_phash, bound_phash) <= 8
                        before_match = before_match or (bound_index == 0 and matched)
                        after_match = after_match or (bound_index == 1 and matched)
                except Exception as exc:
                    add(result, "errors", f"revision page {page} independent evidence comparison failed: {exc}")
                if before_match and after_match:
                    matched_pair = True
                    break
                # A side-by-side before/after image may contain the pair rather
                # than equal either page render.
                errors_before = len(result["errors"])
                validate_contact_sheet_pair(
                    candidate,
                    required[0],
                    required[1],
                    f"revision page {page} evidence",
                    result,
                )
                if len(result["errors"]) == errors_before:
                    matched_pair = True
                    break
                # Suppress the probe-specific messages; emit one deterministic
                # page-level failure below if no candidate proves the pair.
                del result["errors"][errors_before:]
            if not matched_pair:
                add(result, "errors", f"revision page {page} evidence does not bind independent before/after renders")
        if contact_sheet_path is not None:
            validate_contact_sheet_contains(
                contact_sheet_path,
                [*original_render_pages, *revised_render_pages],
                "revision whole-deck contact sheet",
                result,
            )

    original_notes = embedded_notes_texts(original_path, result, "original_pptx notes audit") or []
    revised_notes = embedded_notes_texts(revised_path, result, "revised_pptx notes audit") or []
    checkpoints_by_page: dict[int, list[dict[str, Any]]] = {}
    for checkpoint in pages:
        if isinstance(checkpoint, dict):
            page_number = entry_page_number(checkpoint)
            if page_number is not None:
                checkpoints_by_page.setdefault(page_number, []).append(checkpoint)
    for page in range(1, min(original_count, revised_count) + 1):
        page_checkpoints = checkpoints_by_page.get(page, [])
        original_slide = original_info.slides[page - 1]
        revised_slide = revised_info.slides[page - 1]
        if not page_checkpoints:
            if compact_text(original_slide.visible_text) != compact_text(revised_slide.visible_text):
                add(result, "errors", f"unrequested page {page} native visible text changed")
            if compact_text(original_slide.title) != compact_text(revised_slide.title):
                add(result, "errors", f"unrequested page {page} title/order changed")
            if original_slide.semantic_fingerprint != revised_slide.semantic_fingerprint:
                add(result, "errors", f"unrequested page {page} objects, geometry, media, styles, or relationships changed")
        elif not any(
            checkpoint.get("title_changed") is True
            or normalize_text(checkpoint.get("object_id")).casefold() in {"title", "title-01", "slide-title"}
            for checkpoint in page_checkpoints
        ) and compact_text(original_slide.title) != compact_text(revised_slide.title):
            add(result, "errors", f"revision page {page} changed its title without title_changed=true")
        notes_changed = any(checkpoint.get("notes_changed") is True for checkpoint in page_checkpoints)
        if page <= len(original_notes) and page <= len(revised_notes) and not notes_changed:
            if compact_text(original_notes[page - 1]) != compact_text(revised_notes[page - 1]):
                add(result, "errors", f"revision page {page} changed speaker notes without notes_changed=true")
    if original_info.global_fingerprint != revised_info.global_fingerprint:
        allow_global = data.get("allow_global_changes") is True or data.get("allow_global_change") is True
        if not allow_global:
            add(result, "errors", "deck revision changed masters, layouts, themes, or global presentation settings without explicit approval")
        else:
            validate_bound_evidence_file(
                data.get("global_change_evidence"),
                resolved.parent,
                "revision global change evidence",
                result,
            )
    if data.get("status") not in {None, "passed"}:
        add(result, "errors", "revision-checkpoints overall status must be 'passed'")
    result["metrics"]["revision_checkpoint_pages"] = sorted(changed_pages)


def compare_counts(counts: dict[str, int | None], expected: int | None, result: dict[str, Any]) -> None:
    known = {name: value for name, value in counts.items() if value is not None}
    if expected is not None:
        for name, value in known.items():
            if value != expected:
                add(result, "errors", f"{name} count {value} != expected {expected}")
    if len(set(known.values())) > 1:
        add(result, "errors", f"page/slide counts disagree: {known}")


def enforce_final_contract(args: argparse.Namespace, result: dict[str, Any]) -> None:
    """Require whole-delivery evidence unless the caller explicitly requests a diagnostic smoke test."""
    if args.structural_only:
        add(
            result,
            "warnings",
            "structural-only mode waived final cross-artifact authority/evidence bundles; this result cannot authorize delivery",
        )
        return

    delivery_requested = any(
        (
            args.report,
            args.outline,
            args.style_options,
            args.template_layouts,
            args.template_overview,
            args.slide_images,
            args.image_pptx,
            args.editable_pptx,
            args.speaker_notes,
            args.ppt_notes_pptx,
            args.original_pptx,
            args.revised_pptx,
            args.revision_checkpoints,
        )
    )
    if not delivery_requested:
        add(
            result,
            "errors",
            "final-delivery mode requires an actual deliverable artifact; use --structural-only to inspect ledgers/evidence alone",
        )
    else:
        if not args.project_state:
            add(result, "errors", "final-delivery QA requires --project-state with a passing verify checkpoint")
        if not args.source_manifest:
            add(result, "errors", "final-delivery QA requires top-level --source-manifest")
        if not args.claim_ledger:
            add(result, "errors", "final-delivery QA requires top-level --claim-ledger")

    if args.outline and not args.slide_copy_ledger:
        add(result, "errors", "final outline QA requires --slide-copy-ledger")
    if args.report and not args.report_render_evidence:
        add(result, "errors", "final report QA requires --report-render-evidence")
    if args.slide_images:
        if not args.slide_manifest:
            add(result, "errors", "final slide-images QA requires --slide-manifest")
        if not args.slide_copy_ledger:
            add(result, "errors", "final slide-images QA requires --slide-copy-ledger")
    if args.image_pptx:
        if not args.slide_manifest:
            add(result, "errors", "final image-pptx QA requires --slide-manifest with approved source images")
        if not args.slide_copy_ledger:
            add(result, "errors", "final image-pptx QA requires --slide-copy-ledger")
        if not args.image_pptx_render_evidence:
            add(result, "errors", "final image-pptx QA requires --image-pptx-render-evidence")
    if args.editable_pptx:
        for value, flag in (
            (args.slide_copy_ledger, "--slide-copy-ledger"),
            (args.slide_manifest, "--slide-manifest"),
            (args.editable_handoff, "--editable-handoff"),
            (args.editable_validation, "--editable-validation"),
            (args.offline_preparation, "--offline-preparation"),
        ):
            if not value:
                add(result, "errors", f"final editable-pptx QA requires {flag}")
    if args.template_layouts or args.template_overview or args.template_composer:
        for value, flag in (
            (args.visual_system, "--visual-system"),
            (args.template_layouts, "--template-layouts"),
            (args.template_overview, "--template-overview"),
            (args.template_composer, "--template-composer"),
        ):
            if not value:
                add(result, "errors", f"final template QA requires {flag}")
    if args.speaker_notes:
        for value, flag in (
            (args.speaker_notes_ledger, "--speaker-notes-ledger"),
            (args.slide_copy_ledger, "--slide-copy-ledger"),
            (args.reference_pptx, "--reference-pptx"),
        ):
            if not value:
                add(result, "errors", f"final speaker-notes QA requires {flag}")
        if not args.speaker_notes_render_evidence:
            add(result, "errors", "final speaker-notes QA requires --speaker-notes-render-evidence")
    if args.ppt_notes_pptx:
        for value, flag in (
            (args.speaker_notes_ledger, "--speaker-notes-ledger"),
            (args.slide_copy_ledger, "--slide-copy-ledger"),
            (args.reference_pptx, "--reference-pptx"),
        ):
            if not value:
                add(result, "errors", f"final ppt-notes QA requires {flag}")
    if (args.speaker_notes or args.ppt_notes_pptx) and args.reference_pptx:
        if args.reference_pptx.expanduser().resolve().is_file():
            reference_probe = inspect_opc_package(
                args.reference_pptx.expanduser().resolve(), result, "reference_pptx contract probe", "ppt/presentation.xml"
            )
            if reference_probe is not None:
                has_native_title = False
                for part in ordered_slide_parts(reference_probe, result, "reference_pptx contract probe"):
                    root = reference_probe.roots.get(part)
                    if root is not None and slide_title(root, ppt_paragraphs(root)):
                        has_native_title = True
                        break
                if not has_native_title and not args.slide_manifest:
                    add(result, "errors", "image-only reference PPTX requires --slide-manifest for title/page authority")
    revision_values = (args.original_pptx, args.revised_pptx, args.revision_checkpoints)
    if any(revision_values) and not all(revision_values):
        add(
            result,
            "errors",
            "deck-revision QA requires --original-pptx, --revised-pptx, and --revision-checkpoints together",
        )


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""Machine-ledger JSON examples (schema_version 1):
  slide copy:  {"schema_version":1,"slides":[{"page":1,"id":"slide-01","title":"Title","purpose":"Show the central comparison","visual_type":"comparison","body_copy":["Copy"],"critical_fields":["2026","42 kg"],"source_ids":["src-001"],"footer":"Source: src-001","claim_ids":["claim-001"],"speaker_takeaway":"State the decision implication"}]}
  slide manifest: {"schema_version":1,"slides":[{"page":1,"id":"slide-01","path":"/abs/01.png","sha256":"...","width":1920,"height":1080,"title":"Title","critical_fields":["2026"],"observed_text":["Title","Copy","2026"],"text_verified":true,"text_verification_evidence":"/abs/qa/page-01.json","allow_full_bleed_image":false}]}
  speaker notes: {"schema_version":1,"total_seconds":60,"slides":[{"page":1,"title":"Title","target_seconds":60,"goal":"...","script":"...","cue":"...","transition":"...","claim_ids":["claim-001"]}]}

--editable-validation accepts an editppt run directory, its page_jobs/deck manifest,
or bound standalone JSON whose pages and final_validation reference hashed passed JSON files.
""",
    )
    parser.add_argument("--slide-images", type=Path, help="Directory containing ordered PNG/JPEG slide images")
    parser.add_argument("--image-pptx", type=Path, help="Image-based PPTX to validate")
    parser.add_argument("--editable-pptx", type=Path, help="Object-level editable PPTX to validate")
    parser.add_argument("--report", type=Path, help="Report DOCX")
    parser.add_argument("--report-render-evidence", type=Path, help="Hash-bound ppt-gen.render-evidence JSON for ordered report page renders")
    parser.add_argument("--outline", type=Path, help="Outline DOCX, Markdown, text, or JSON")
    parser.add_argument("--speaker-notes", type=Path, help="Speaker-notes DOCX with one '第 N 页'/'Slide N' section per slide")
    parser.add_argument("--speaker-notes-render-evidence", type=Path, help="Hash-bound ppt-gen.render-evidence JSON for ordered notes page renders")
    parser.add_argument("--image-pptx-render-evidence", type=Path, help="Hash-bound ppt-gen.render-evidence JSON for ordered image-PPTX slide renders")
    parser.add_argument("--ppt-notes-pptx", type=Path, help="PPTX whose native embedded notesSlides must match the speaker-notes ledger")
    parser.add_argument("--reference-pptx", type=Path, help="Accepted reference deck used to prove speaker-note page/title mapping")
    parser.add_argument("--original-pptx", type=Path, help="Immutable original PPTX for a deck-revision delivery")
    parser.add_argument("--revised-pptx", type=Path, help="Versioned revised PPTX for a deck-revision delivery")
    parser.add_argument("--revision-checkpoints", type=Path, help="Revision authority JSON binding original/revised hashes, requested changes, and visual evidence")
    parser.add_argument("--expected-slides", type=positive_int, help="Required positive slide/page count")
    parser.add_argument("--json-out", type=Path, help="Optional path for the JSON result")
    parser.add_argument("--cache-dir", type=Path, help="Optional project-private OCR/render cache; default is auto-cleaned process-local storage")
    parser.add_argument("--slide-copy-ledger", type=Path, help="Contract JSON with page, title, body_copy, and critical_fields (legacy aliases accepted)")
    parser.add_argument("--slide-manifest", type=Path, help="Contract JSON with page/path/SHA-256/dimensions/title/critical_fields")
    parser.add_argument("--speaker-notes-ledger", type=Path, help="Contract JSON with total_seconds and per-page target_seconds/script/cues")
    parser.add_argument("--source-manifest", type=Path, help="Top-level source authority JSON with policies, paths, and current hashes")
    parser.add_argument("--claim-ledger", type=Path, help="Top-level factual authority JSON; pending/unsupported claims fail final QA")
    parser.add_argument("--project-state", type=Path, help="Required in final-delivery mode; binds policies, a fresh verify checkpoint, style, and registered artifacts")
    parser.add_argument("--editable-validation", type=Path, help="editppt run directory or JSON validation evidence bound to the editable PPTX")
    parser.add_argument("--editable-handoff", type=Path, help="editable-handoff.json binding offline/privacy policy, ledgers, and source images")
    parser.add_argument("--offline-preparation", type=Path, help="Hash-bound guarded prepare and actual offline OCR execution record")
    parser.add_argument("--structural-only", action="store_true", help="Diagnostic smoke-test mode: waive final cross-artifact bundles; never use for completion or delivery")
    parser.add_argument("--style-options", type=Path, help="style-options.json with 2-3 hashed exact-16:9 candidates and selected_id")
    parser.add_argument("--visual-system", type=Path, help="visual-system.json with palette, typography, geometry, and six layouts")
    parser.add_argument("--template-layouts", type=Path, help="Directory containing exactly six exact-16:9 layouts; numbering is required only without composer inputs")
    parser.add_argument("--template-overview", type=Path, help="Exact 1920x1080 template overview PNG")
    parser.add_argument("--template-composer", type=Path, help="Optional composer JSON whose thumbnail_size and overview hash are checked")
    parser.add_argument("--min-image-width", type=positive_int, default=1, help="Optional minimum decoded slide-image width (default: 1 for legacy compatibility)")
    parser.add_argument("--min-image-height", type=positive_int, default=1, help="Optional minimum decoded slide-image height (default: 1 for legacy compatibility)")
    parser.add_argument(
        "--full-slide-image-threshold",
        type=unit_interval,
        default=0.90,
        help="Fraction of slide area treated as a full-slide raster (default: 0.90)",
    )
    parser.add_argument(
        "--duration-tolerance-percent",
        type=percentage,
        default=10.0,
        help="Maximum speaker-note duration drift before failure (default: 10)",
    )
    return parser


def finalize_status(result: dict[str, Any]) -> None:
    if result["errors"]:
        result["status"] = "failed"
        result["passed"] = False
    elif result["warnings"]:
        result["status"] = "passed_with_warnings"
        result["passed"] = True
    else:
        result["status"] = "passed"
        result["passed"] = True


def sha256_directory(path: Path) -> str:
    """Match project_state.file_record's canonical directory-v1 digest."""
    digest = hashlib.sha256()
    digest.update(b"ppt-gen-directory-v1\0")
    digest.update(b"D\0.\0")
    for current, dirnames, filenames in os.walk(path, followlinks=False):
        current_path = Path(current)
        dirnames.sort()
        filenames.sort()
        for dirname in list(dirnames):
            child = current_path / dirname
            relative = child.relative_to(path).as_posix()
            if child.is_symlink():
                digest.update(f"L\0{relative}\0{os.readlink(child)}\0".encode("utf-8"))
                dirnames.remove(dirname)
            else:
                digest.update(f"D\0{relative}\0".encode("utf-8"))
        for filename in filenames:
            child = current_path / filename
            relative = child.relative_to(path).as_posix()
            stat = child.lstat()
            if child.is_symlink():
                digest.update(f"L\0{relative}\0{os.readlink(child)}\0".encode("utf-8"))
                continue
            if not child.is_file():
                digest.update(f"O\0{relative}\0{stat.st_mode}\0".encode("utf-8"))
                continue
            digest.update(f"F\0{relative}\0".encode("utf-8"))
            with child.open("rb") as handle:
                for chunk in iter(lambda: handle.read(1024 * 1024), b""):
                    digest.update(chunk)
    return digest.hexdigest()


def collect_artifacts(args: argparse.Namespace, result: dict[str, Any]) -> list[dict[str, Any]]:
    roles = (
        ("slide-images", args.slide_images),
        ("image-pptx", args.image_pptx),
        ("editable-pptx", args.editable_pptx),
        ("report", args.report),
        ("report-render-evidence", args.report_render_evidence),
        ("outline", args.outline),
        ("speaker-notes", args.speaker_notes),
        ("speaker-notes-render-evidence", args.speaker_notes_render_evidence),
        ("image-pptx-render-evidence", args.image_pptx_render_evidence),
        ("ppt-notes-pptx", args.ppt_notes_pptx),
        ("reference-pptx", args.reference_pptx),
        ("original-pptx", args.original_pptx),
        ("revised-pptx", args.revised_pptx),
        ("revision-checkpoints", args.revision_checkpoints),
        ("source-manifest", args.source_manifest),
        ("claim-ledger", args.claim_ledger),
        ("slide-copy-ledger", args.slide_copy_ledger),
        ("slide-manifest", args.slide_manifest),
        ("speaker-notes-ledger", args.speaker_notes_ledger),
        ("editable-validation", args.editable_validation),
        ("editable-handoff", args.editable_handoff),
        ("offline-preparation", args.offline_preparation),
        ("style-options", args.style_options),
        ("visual-system", args.visual_system),
        ("template-layouts", args.template_layouts),
        ("template-overview", args.template_overview),
        ("template-composer", args.template_composer),
    )
    artifacts: list[dict[str, Any]] = []
    for role, value in roles:
        if value is None:
            continue
        path = value.expanduser().resolve()
        entry: dict[str, Any] = {"role": role, "path": str(path), "sha256": None}
        try:
            if path.is_file():
                entry["sha256"] = sha256_file(path)
            elif path.is_dir():
                entry["sha256"] = sha256_directory(path)
                entry["hash_scope"] = "recursive-files"
        except OSError as exc:
            add(result, "errors", f"cannot hash artifact {role} at {path}: {exc}")
        artifacts.append(entry)
    return artifacts


def build_gates(args: argparse.Namespace, result: dict[str, Any]) -> list[dict[str, Any]]:
    configured = [
        ("images", bool(args.slide_images or args.slide_manifest)),
        ("pptx.image", bool(args.image_pptx)),
        ("pptx.editable", bool(args.editable_pptx or args.editable_validation)),
        ("docx.report", bool(args.report)),
        ("outline", bool(args.outline)),
        ("speaker-notes", bool(args.speaker_notes or args.speaker_notes_ledger)),
        ("speaker-notes.embedded", bool(args.ppt_notes_pptx)),
        ("deck-revision", bool(args.original_pptx or args.revised_pptx or args.revision_checkpoints)),
        ("authority.sources", bool(args.source_manifest)),
        ("authority.claims", bool(args.claim_ledger)),
        ("state", bool(args.project_state)),
        ("style-options", bool(args.style_options)),
        ("visual-system", bool(args.visual_system)),
        ("template", bool(args.template_layouts or args.template_overview)),
        ("cross-artifact", True),
    ]
    overall = result["status"]
    evidence = [artifact["path"] for artifact in result.get("artifacts", []) if artifact["sha256"]]
    messages = list(result["errors"] if result["errors"] else result["warnings"])
    return [
        {
            "id": gate_id,
            "status": overall,
            "evidence": evidence,
            "messages": messages,
        }
        for gate_id, enabled in configured
        if enabled
    ]


def main() -> None:
    args = build_parser().parse_args()
    configure_cache(args.cache_dir)
    artifacts = (
        args.slide_images,
        args.image_pptx,
        args.editable_pptx,
        args.report,
        args.report_render_evidence,
        args.outline,
        args.speaker_notes,
        args.speaker_notes_render_evidence,
        args.image_pptx_render_evidence,
        args.ppt_notes_pptx,
        args.reference_pptx,
        args.original_pptx,
        args.revised_pptx,
        args.revision_checkpoints,
        args.editable_validation,
        args.editable_handoff,
        args.offline_preparation,
        args.source_manifest,
        args.claim_ledger,
        args.project_state,
        args.style_options,
        args.visual_system,
        args.template_layouts,
        args.template_overview,
        args.slide_copy_ledger,
        args.slide_manifest,
        args.speaker_notes_ledger,
        args.template_composer,
    )
    if not any(artifacts):
        raise SystemExit("Provide at least one artifact to validate")

    result: dict[str, Any] = {
        "schema_version": 1,
        "tool": "ppt-gen.validate_delivery",
        "validation_contract_version": 2,
        "mode": "structural-only" if args.structural_only else "final-delivery",
        "checked_at": dt.datetime.now(dt.timezone.utc).astimezone().isoformat(timespec="seconds"),
        "status": "failed",
        "passed": False,
        "artifacts": [],
        "gates": [],
        "errors": [],
        "warnings": [],
        "metrics": {},
    }
    counts: dict[str, int | None] = {}

    enforce_final_contract(args, result)
    if args.offline_preparation:
        result["errors"].extend(validate_offline_preparation(args.offline_preparation, args.editable_handoff))
    source_data = validate_source_manifest(args.source_manifest.expanduser().resolve(), result) if args.source_manifest else None
    claim_data = (
        validate_claim_ledger(
            args.claim_ledger,
            result,
            source_data,
        )
        if args.claim_ledger
        else None
    )
    state_data = (
        validate_project_state(args.project_state, result, source_data, args.style_options)
        if args.project_state
        else None
    )

    _, slide_copy = load_ledger(args.slide_copy_ledger, "slide-copy ledger", result)
    slide_manifest_data, slide_manifest = load_ledger(args.slide_manifest, "slide manifest", result)
    notes_data, notes_ledger = load_ledger(args.speaker_notes_ledger, "speaker-notes ledger", result)
    if slide_copy is not None:
        counts["slide_copy_ledger"] = len(slide_copy)
    if slide_manifest is not None:
        counts["slide_manifest"] = len(slide_manifest)
    if notes_ledger is not None:
        counts["speaker_notes_ledger"] = len(notes_ledger)
    active_data_handling = (
        (state_data.get("policies") or {}).get("data_handling")
        if isinstance(state_data, dict) and isinstance(state_data.get("policies"), dict)
        else source_data.get("data_handling") if isinstance(source_data, dict) else None
    )
    if slide_manifest_data is not None:
        validate_visual_generation_provenance(
            slide_manifest_data,
            "slide manifest",
            active_data_handling,
            result,
        )
    if args.speaker_notes or args.ppt_notes_pptx:
        validate_speaker_notes_authorities(
            notes_data,
            args.speaker_notes_ledger,
            args.slide_copy_ledger,
            args.reference_pptx,
            result,
        )
    validate_claim_references(claim_data, slide_copy, notes_ledger, result)

    effective_min_width = args.min_image_width if args.structural_only else max(1280, args.min_image_width)
    effective_min_height = args.min_image_height if args.structural_only else max(720, args.min_image_height)

    slide_images: list[ImageRecord] | None = None
    if args.slide_images:
        slide_images = validate_slide_images(
            args.slide_images.expanduser().resolve(),
            result,
            slide_manifest,
            effective_min_width,
            effective_min_height,
        )
        counts["slide_images"] = len(slide_images) if slide_images is not None else None
    elif args.slide_manifest and slide_manifest is not None:
        slide_images = validate_manifest_images(
            args.slide_manifest,
            slide_manifest,
            result,
            effective_min_width,
            effective_min_height,
        )
    compare_image_manifest_copy(
        slide_images,
        slide_manifest,
        slide_copy,
        args.slide_manifest.expanduser().resolve().parent if args.slide_manifest else None,
        result,
    )

    image_info: PptxInfo | None = None
    if args.image_pptx:
        image_info = pptx_info(
            args.image_pptx.expanduser().resolve(),
            result,
            "image_pptx",
            editable=False,
            ledger=None,
            page_manifest=slide_manifest,
            source_images=slide_images,
            full_slide_threshold=args.full_slide_image_threshold,
        )
        counts["image_pptx"] = len(image_info.slides) if image_info is not None else None
    if args.image_pptx_render_evidence:
        validate_render_evidence_path(
            args.image_pptx_render_evidence,
            args.image_pptx,
            "image_pptx render evidence",
            result,
            len(image_info.slides) if image_info is not None else args.expected_slides,
        )

    editable_info: PptxInfo | None = None
    editable_path = args.editable_pptx.expanduser().resolve() if args.editable_pptx else None
    if editable_path:
        editable_info = pptx_info(
            editable_path,
            result,
            "editable_pptx",
            editable=True,
            ledger=slide_copy,
            page_manifest=slide_manifest,
            source_images=slide_images,
            full_slide_threshold=args.full_slide_image_threshold,
        )
        counts["editable_pptx"] = len(editable_info.slides) if editable_info is not None else None
        if not args.structural_only:
            if not args.editable_validation:
                add(result, "errors", "--editable-pptx requires --editable-validation with per-page and final companion evidence")
            if not args.editable_handoff:
                add(result, "errors", "--editable-pptx requires --editable-handoff for OCR/privacy and source-authority policy")
            if slide_copy is None:
                add(result, "errors", "--editable-pptx requires --slide-copy-ledger for authoritative page-text reconciliation")
            if slide_manifest is None or slide_images is None:
                add(result, "errors", "--editable-pptx requires a readable --slide-manifest (and optionally --slide-images) for source identity")
        else:
            add(result, "warnings", "structural-only mode waived editable ledger/evidence gates; result is not final-delivery QA")

    if args.editable_validation:
        counts["editable_validation"] = validate_editppt_evidence(
            args.editable_validation,
            result,
            editable_path,
            slide_images,
        )
    if args.editable_handoff:
        validate_editable_handoff(
            args.editable_handoff,
            result,
            args.source_manifest,
            args.slide_copy_ledger,
            args.slide_manifest,
            slide_images,
        )

    if args.report:
        report_info = validate_docx(args.report.expanduser().resolve(), result, "report")
        validate_report_claim_coverage(report_info, claim_data, result)
    if args.report_render_evidence:
        validate_render_evidence_path(
            args.report_render_evidence,
            args.report,
            "report render evidence",
            result,
        )
    if args.outline:
        outline = args.outline.expanduser().resolve()
        outline_text: str | None = None
        if outline.suffix.lower() == ".docx":
            outline_info = validate_docx(outline, result, "outline")
            outline_text = outline_info.text if outline_info is not None else None
        else:
            outline_text = validate_text_outline(outline, result)
        reconcile_outline_to_slide_copy(outline_text, slide_copy, result)

    if args.style_options:
        targets = state_data.get("targets", []) if isinstance(state_data, dict) else []
        downstream_visual_targets = {"template", "image-deck", "editable-deck"}
        require_style_selection = bool(
            args.visual_system
            or args.template_layouts
            or args.template_overview
            or args.slide_images
            or args.image_pptx
            or args.editable_pptx
            or any(target in downstream_visual_targets for target in targets)
        )
        validate_style_options(
            args.style_options,
            result,
            require_style_selection,
            active_data_handling,
        )
    if args.visual_system:
        validate_visual_system(
            args.visual_system,
            result,
            args.style_options,
            active_data_handling,
        )
    if args.template_layouts:
        validate_template_layouts(
            args.template_layouts,
            result,
            effective_min_width,
            effective_min_height,
            args.template_composer,
        )
    if args.template_overview:
        validate_template_overview(
            args.template_overview, args.template_composer, args.visual_system, result
        )
    elif args.template_composer:
        add(result, "errors", "--template-composer requires --template-overview")

    reference_info: PptxInfo | None = None
    if args.reference_pptx:
        reference_info = pptx_info(
            args.reference_pptx.expanduser().resolve(),
            result,
            "reference_pptx",
            editable=False,
            ledger=None,
            page_manifest=slide_manifest,
            source_images=slide_images,
            full_slide_threshold=args.full_slide_image_threshold,
        )
        counts["reference_pptx"] = len(reference_info.slides) if reference_info is not None else None
        if reference_info is not None and slide_copy is not None:
            if len(reference_info.slides) != len(slide_copy):
                add(result, "errors", "reference_pptx page count does not match slide-copy ledger")
            else:
                for slide, entry in zip(reference_info.slides, slide_copy):
                    expected_title = normalize_text(entry.get("title", ""))
                    if expected_title and slide.title:
                        if compact_text(slide.title) != compact_text(expected_title):
                            add(
                                result,
                                "errors",
                                f"reference_pptx slide {slide.index} title mismatch: {slide.title!r} != {expected_title!r}",
                            )
                    elif expected_title:
                        manifest_title = normalize_text(
                            slide_manifest[slide.index - 1].get("title", "")
                            if slide_manifest and slide.index <= len(slide_manifest)
                            else ""
                        )
                        if compact_text(manifest_title) != compact_text(expected_title):
                            add(
                                result,
                                "errors",
                                f"reference_pptx image-only slide {slide.index} requires a matching --slide-manifest title",
                            )

    ppt_notes_info: PptxInfo | None = None
    if args.ppt_notes_pptx:
        ppt_notes_path = args.ppt_notes_pptx.expanduser().resolve()
        ppt_notes_info = pptx_info(
            ppt_notes_path,
            result,
            "ppt_notes_pptx",
            editable=False,
            ledger=None,
            page_manifest=slide_manifest,
            source_images=slide_images,
            full_slide_threshold=args.full_slide_image_threshold,
        )
        counts["ppt_notes_pptx"] = len(ppt_notes_info.slides) if ppt_notes_info is not None else None
        validate_embedded_notes_pptx(
            ppt_notes_path,
            ppt_notes_info,
            notes_ledger,
            slide_copy,
            slide_manifest,
            result,
        )

    if args.original_pptx and args.revised_pptx and args.revision_checkpoints:
        original_path = args.original_pptx.expanduser().resolve()
        revised_path = args.revised_pptx.expanduser().resolve()
        original_info = pptx_info(
            original_path,
            result,
            "original_pptx",
            editable=False,
            ledger=None,
            page_manifest=None,
            source_images=None,
            full_slide_threshold=args.full_slide_image_threshold,
        )
        revised_info = pptx_info(
            revised_path,
            result,
            "revised_pptx",
            editable=False,
            ledger=None,
            page_manifest=None,
            source_images=None,
            full_slide_threshold=args.full_slide_image_threshold,
        )
        validate_revision_checkpoints(
            args.revision_checkpoints,
            original_path,
            revised_path,
            original_info,
            revised_info,
            claim_data,
            result,
        )

    if args.speaker_notes:
        notes_info = validate_docx(args.speaker_notes.expanduser().resolve(), result, "speaker_notes")
        inferred_count = args.expected_slides
        if inferred_count is None:
            for key in ("editable_pptx", "image_pptx", "slide_images", "slide_copy_ledger"):
                if counts.get(key) is not None:
                    inferred_count = counts[key]
                    break
        deck_titles = [slide.title for slide in reference_info.slides] if reference_info is not None else None
        counts["speaker_notes"] = validate_speaker_notes(
            notes_info,
            result,
            inferred_count,
            notes_data,
            notes_ledger,
            slide_copy,
            deck_titles,
            args.duration_tolerance_percent,
        )
    if args.speaker_notes_render_evidence:
        validate_render_evidence_path(
            args.speaker_notes_render_evidence,
            args.speaker_notes,
            "speaker_notes render evidence",
            result,
            # DOCX render pages do not equal slide sections: the document may
            # include a cover, place multiple short sections on one page, or
            # flow one long section across pages. Independent render QA owns
            # physical page count; the ledger/deck checks own section count.
            None,
        )

    if args.speaker_notes or args.ppt_notes_pptx:
        requested_duration = (state_data or {}).get("preferences", {}).get("requested_duration_seconds")
        validate_timing(notes_data, notes_ledger, result, requested_duration, args.duration_tolerance_percent)
    compare_counts(counts, args.expected_slides, result)
    finalize_status(result)
    result["artifacts"] = collect_artifacts(args, result)
    validate_state_artifact_linkage(state_data, result["artifacts"], result)
    result["authorities"] = {
        artifact["role"]: {"path": artifact["path"], "sha256": artifact["sha256"]}
        for artifact in result["artifacts"]
        if artifact["role"] in {"source-manifest", "claim-ledger"}
    }
    if state_data is not None:
        verified = [
            entry
            for entry in state_data.get("history", [])
            if isinstance(entry, dict) and entry.get("event") == "verified"
        ]
        result["state_checkpoint"] = {
            "path": str(args.project_state.expanduser().resolve()) if args.project_state else None,
            "state_revision": state_data.get("state_revision"),
            "verify_history_id": verified[-1].get("id") if verified else None,
            "verify_revision": verified[-1].get("revision") if verified else None,
            "drift_count": (verified[-1].get("details") or {}).get("drift_count") if verified else None,
        }
    finalize_status(result)
    result["gates"] = build_gates(args, result)

    if args.json_out:
        output_path = args.json_out.expanduser().resolve()
        try:
            output_path.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        except OSError as exc:
            add(result, "errors", f"cannot write --json-out {output_path}: {exc}")
            finalize_status(result)
            result["gates"] = build_gates(args, result)

    print(json.dumps(result, ensure_ascii=False, indent=2))
    sys.exit(0 if result["passed"] else 1)


if __name__ == "__main__":
    main()
