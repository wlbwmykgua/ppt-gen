#!/usr/bin/env python3
"""Compose six exact-16:9 slide previews into a deterministic 1920x1080 overview.

The 3x2 geometry is created locally with Pillow. The script never asks an image
generation model to draw or preserve thumbnail geometry.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import sys
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable


CANVAS_SIZE = (1920, 1080)
THUMB_SIZE = (560, 315)
GRID_COLUMNS = 3
GRID_ROWS = 2
GRID_GAP_X = 28
GRID_GAP_Y = 18
GRID_TOP = 108
LABEL_HEIGHT = 30
PANEL_BOX = (56, 838, 1864, 1040)
DEFAULT_LABELS = ("封面", "章节页", "叙事内容", "对比页", "数据图表", "结论页")


def now() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def atomic_write_json(path: Path, payload: dict[str, Any]) -> None:
    path = path.expanduser().resolve()
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_suffix(path.suffix + ".tmp")
    temp.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    os.replace(temp, path)


def normalize_key(value: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", value.casefold())


def get_mapping(value: Any) -> dict[str, Any]:
    return value if isinstance(value, dict) else {}


def lookup(mapping: dict[str, Any], keys: Iterable[str], default: Any = None) -> Any:
    normalized = {normalize_key(str(key)): value for key, value in mapping.items()}
    for key in keys:
        candidate = normalized.get(normalize_key(key))
        if candidate not in (None, ""):
            return candidate
    return default


def parse_color(value: Any, default: str) -> tuple[int, int, int]:
    if isinstance(value, (list, tuple)) and len(value) >= 3:
        channels = tuple(max(0, min(255, int(channel))) for channel in value[:3])
        return channels  # type: ignore[return-value]
    raw = str(value or default).strip()
    if raw.startswith("#"):
        raw = raw[1:]
    if len(raw) == 3 and re.fullmatch(r"[0-9a-fA-F]{3}", raw):
        raw = "".join(character * 2 for character in raw)
    if not re.fullmatch(r"[0-9a-fA-F]{6}", raw):
        raw = default.lstrip("#")
    return tuple(int(raw[index : index + 2], 16) for index in (0, 2, 4))  # type: ignore[return-value]


def color_hex(color: tuple[int, int, int]) -> str:
    return "#" + "".join(f"{channel:02X}" for channel in color)


def luminance(color: tuple[int, int, int]) -> float:
    return 0.2126 * color[0] + 0.7152 * color[1] + 0.0722 * color[2]


def contrast_text(color: tuple[int, int, int]) -> tuple[int, int, int]:
    return (20, 24, 32) if luminance(color) > 155 else (248, 250, 252)


def blend(left: tuple[int, int, int], right: tuple[int, int, int], amount: float) -> tuple[int, int, int]:
    return tuple(round(a * (1 - amount) + b * amount) for a, b in zip(left, right))  # type: ignore[return-value]


def parse_slide(value: str, index: int) -> tuple[str, Path]:
    if "=" in value:
        label, raw_path = value.split("=", 1)
        label = label.strip() or DEFAULT_LABELS[index]
    else:
        raw_path = value
        label = DEFAULT_LABELS[index]
    path = Path(raw_path.strip()).expanduser().resolve()
    return label, path


def visual_tokens(payload: dict[str, Any]) -> dict[str, Any]:
    colors = get_mapping(payload.get("colors") or payload.get("palette"))
    typography = get_mapping(payload.get("typography") or payload.get("fonts"))
    grid = get_mapping(payload.get("grid"))
    spacing = get_mapping(payload.get("spacing"))
    chart = get_mapping(payload.get("chart") or payload.get("charts"))
    imagery = get_mapping(payload.get("imagery") or payload.get("images"))

    primary = parse_color(
        lookup(colors, ("primary", "primary_color", "main", "主色"), payload.get("primary_color")),
        "#2446A8",
    )
    secondary = parse_color(
        lookup(colors, ("secondary", "secondary_color", "辅色"), payload.get("secondary_color")),
        "#5B8DEF",
    )
    accent = parse_color(
        lookup(colors, ("accent", "accent_color", "强调色"), payload.get("accent_color")),
        "#F2A541",
    )
    background = parse_color(
        lookup(colors, ("background", "background_color", "背景色"), payload.get("background_color")),
        "#F4F6FA",
    )
    title_font = str(
        lookup(typography, ("title_font", "heading_font", "display", "标题字体"), payload.get("title_font") or "PingFang SC")
    )
    body_font = str(
        lookup(typography, ("body_font", "body", "正文", "正文字体"), payload.get("body_font") or "PingFang SC")
    )
    return {
        "name": str(payload.get("name") or payload.get("style_name") or payload.get("title") or "已选视觉系统"),
        "primary": primary,
        "secondary": secondary,
        "accent": accent,
        "background": background,
        "title_font": title_font,
        "body_font": body_font,
        "grid": str(
            lookup(grid, ("summary", "description", "columns", "网格"), payload.get("grid_summary") or "12-column grid")
        ),
        "spacing": str(
            lookup(spacing, ("summary", "description", "base", "spacing_unit", "间距"), payload.get("spacing_summary") or "8px rhythm")
        ),
        "chart": str(
            lookup(chart, ("summary", "description", "style", "图表"), payload.get("chart_summary") or "Direct labels; restrained gridlines")
        ),
        "imagery": str(
            lookup(imagery, ("summary", "description", "treatment", "style", "图片"), payload.get("imagery_summary") or "Consistent crop and color treatment")
        ),
        "render_font_path": lookup(typography, ("render_font_path", "font_path"), payload.get("render_font_path")),
    }


def font_candidates(explicit: Path | None, tokens: dict[str, Any]) -> list[Path]:
    candidates = []
    if explicit:
        candidates.append(explicit.expanduser())
    if tokens.get("render_font_path"):
        candidates.append(Path(str(tokens["render_font_path"])).expanduser())
    candidates.extend(
        [
            Path("/System/Library/Fonts/PingFang.ttc"),
            Path("/System/Library/Fonts/Hiragino Sans GB.ttc"),
            Path("/System/Library/Fonts/STHeiti Medium.ttc"),
            Path("/System/Library/Fonts/STHeiti Light.ttc"),
            Path("/System/Library/Fonts/Supplemental/Songti.ttc"),
            Path("/System/Library/Fonts/Supplemental/Arial Unicode.ttf"),
            Path("/System/Library/Fonts/Helvetica.ttc"),
            Path("/Library/Fonts/Arial Unicode.ttf"),
            Path("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"),
            Path("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"),
        ]
    )
    return candidates


def load_fonts(explicit: Path | None, tokens: dict[str, Any]) -> tuple[dict[str, Any], str]:
    from PIL import ImageFont

    selected = next((path for path in font_candidates(explicit, tokens) if path.is_file()), None)
    if selected:
        return {
            "title": ImageFont.truetype(str(selected), 34),
            "label": ImageFont.truetype(str(selected), 18),
            "panel_title": ImageFont.truetype(str(selected), 19),
            "panel_body": ImageFont.truetype(str(selected), 15),
            "swatch": ImageFont.truetype(str(selected), 13),
        }, str(selected.resolve())
    default = ImageFont.load_default()
    return {key: default for key in ("title", "label", "panel_title", "panel_body", "swatch")}, "Pillow default"


def truncate(draw: Any, text: str, font: Any, max_width: int) -> str:
    text = " ".join(str(text).split())
    if draw.textbbox((0, 0), text, font=font)[2] <= max_width:
        return text
    suffix = "…"
    while text and draw.textbbox((0, 0), text + suffix, font=font)[2] > max_width:
        text = text[:-1]
    return text.rstrip() + suffix


def draw_swatch(draw: Any, x: int, y: int, label: str, color: tuple[int, int, int], fonts: dict[str, Any]) -> None:
    draw.rounded_rectangle((x, y, x + 78, y + 48), radius=8, fill=color, outline=blend(color, (0, 0, 0), 0.18), width=1)
    draw.text((x + 8, y + 8), label, font=fonts["swatch"], fill=contrast_text(color))
    draw.text((x + 8, y + 27), color_hex(color), font=fonts["swatch"], fill=contrast_text(color))


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Validate six exact-16:9 slides and compose a deterministic 1920x1080 3x2 overview.",
        epilog="Use --slide label=/absolute/path.png exactly six times.",
    )
    parser.add_argument("--slide", action="append", required=True, metavar="label=path", help="Slide preview; repeat exactly six times.")
    parser.add_argument("--visual-system", type=Path, required=True, help="JSON containing colors, typography, grid, chart, imagery, and spacing tokens.")
    parser.add_argument("--output", type=Path, required=True, help="Output PNG path.")
    parser.add_argument("--title", help="Optional overview title override.")
    parser.add_argument("--font-path", type=Path, help="Optional local font file used to render overview labels.")
    parser.add_argument("--json-out", type=Path, help="Optional JSON result path.")
    parser.add_argument("--force", action="store_true", help="Replace an existing output image.")
    return parser


def run(args: argparse.Namespace) -> dict[str, Any]:
    try:
        from PIL import Image, ImageDraw
    except ImportError as exc:
        raise RuntimeError("Pillow is required: install the workspace-provided Pillow runtime before composing the overview.") from exc

    if len(args.slide) != 6:
        raise ValueError(f"Exactly six --slide inputs are required; received {len(args.slide)}")
    visual_path = args.visual_system.expanduser().resolve()
    if not visual_path.is_file():
        raise ValueError(f"Visual-system JSON is missing: {visual_path}")
    try:
        visual_payload = json.loads(visual_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise ValueError(f"Cannot read visual-system JSON: {exc}") from exc
    if not isinstance(visual_payload, dict):
        raise ValueError("Visual-system JSON must contain an object at the top level")
    tokens = visual_tokens(visual_payload)

    parsed = [parse_slide(value, index) for index, value in enumerate(args.slide)]
    inputs = []
    opened = []
    for label, path in parsed:
        if not path.is_file():
            raise ValueError(f"Slide preview is missing: {path}")
        try:
            image = Image.open(path)
            image.load()
        except OSError as exc:
            raise ValueError(f"Cannot open slide preview {path}: {exc}") from exc
        width, height = image.size
        if width * 9 != height * 16:
            image.close()
            raise ValueError(f"Slide preview is not exact 16:9: {path} ({width}x{height})")
        inputs.append(
            {
                "label": label,
                "path": str(path),
                "sha256": sha256_file(path),
                "width": width,
                "height": height,
                "exact_16_9": True,
            }
        )
        opened.append(image.convert("RGB"))
        image.close()

    output_path = args.output.expanduser().resolve()
    if output_path.suffix.casefold() != ".png":
        raise ValueError("The overview output must use a .png extension")
    if output_path.exists() and not args.force:
        raise ValueError(f"Output already exists; use --force to replace it: {output_path}")
    output_path.parent.mkdir(parents=True, exist_ok=True)

    background = tokens["background"]
    canvas = Image.new("RGB", CANVAS_SIZE, color=background)
    draw = ImageDraw.Draw(canvas)
    fonts, render_font = load_fonts(args.font_path, tokens)
    ink = contrast_text(background)
    muted = blend(ink, background, 0.42)
    border = blend(ink, background, 0.78)
    panel_fill = blend(background, (255, 255, 255) if luminance(background) < 150 else (0, 0, 0), 0.055)

    title = args.title or f"模板总览 · {tokens['name']}"
    draw.rounded_rectangle((56, 28, 72, 78), radius=8, fill=tokens["primary"])
    draw.text((90, 30), truncate(draw, title, fonts["title"], 1350), font=fonts["title"], fill=ink)
    draw.text((1610, 40), "1920×1080 · 3×2", font=fonts["label"], fill=muted)

    grid_width = GRID_COLUMNS * THUMB_SIZE[0] + (GRID_COLUMNS - 1) * GRID_GAP_X
    grid_left = (CANVAS_SIZE[0] - grid_width) // 2
    resized_filter = getattr(Image, "Resampling", Image).LANCZOS
    for index, ((label, _), image) in enumerate(zip(parsed, opened)):
        row, column = divmod(index, GRID_COLUMNS)
        x = grid_left + column * (THUMB_SIZE[0] + GRID_GAP_X)
        y = GRID_TOP + row * (THUMB_SIZE[1] + LABEL_HEIGHT + GRID_GAP_Y)
        thumbnail = image.resize(THUMB_SIZE, resized_filter)
        canvas.paste(thumbnail, (x, y))
        draw.rectangle((x, y, x + THUMB_SIZE[0] - 1, y + THUMB_SIZE[1] - 1), outline=border, width=2)
        label_text = f"{index + 1:02d}  {label}"
        draw.text((x + 2, y + THUMB_SIZE[1] + 5), truncate(draw, label_text, fonts["label"], THUMB_SIZE[0] - 4), font=fonts["label"], fill=ink)
        image.close()

    draw.rounded_rectangle(PANEL_BOX, radius=18, fill=panel_fill, outline=border, width=2)
    panel_y = PANEL_BOX[1] + 18
    draw.text((82, panel_y), "色彩", font=fonts["panel_title"], fill=ink)
    swatch_y = panel_y + 34
    for offset, (label, color) in enumerate(
        (("主色", tokens["primary"]), ("辅色", tokens["secondary"]), ("强调", tokens["accent"]), ("背景", tokens["background"]))
    ):
        draw_swatch(draw, 82 + offset * 92, swatch_y, label, color, fonts)

    column_specs = [
        (500, "字体", f"标题：{tokens['title_font']}\n正文：{tokens['body_font']}", 390),
        (930, "网格与间距", f"网格：{tokens['grid']}\n间距：{tokens['spacing']}", 360),
        (1328, "图表与图片", f"图表：{tokens['chart']}\n图片：{tokens['imagery']}", 470),
    ]
    for x, heading, body, width in column_specs:
        draw.text((x, panel_y), heading, font=fonts["panel_title"], fill=ink)
        lines = body.splitlines()
        for line_index, line in enumerate(lines):
            draw.text(
                (x, panel_y + 38 + line_index * 31),
                truncate(draw, line, fonts["panel_body"], width),
                font=fonts["panel_body"],
                fill=muted,
            )

    temp_path = output_path.with_suffix(output_path.suffix + ".tmp")
    canvas.save(temp_path, format="PNG", compress_level=9)
    canvas.close()
    os.replace(temp_path, output_path)

    with Image.open(output_path) as verification:
        if verification.size != CANVAS_SIZE:
            raise RuntimeError(f"Output verification failed: {verification.size} != {CANVAS_SIZE}")

    return {
        "schema_version": 1,
        "tool": "ppt-gen-compose-template-overview",
        "generated_at": now(),
        "status": "passed",
        "output": {
            "path": str(output_path),
            "sha256": sha256_file(output_path),
            "width": CANVAS_SIZE[0],
            "height": CANVAS_SIZE[1],
            "layout": "3x2",
            "thumbnail_size": list(THUMB_SIZE),
        },
        "visual_system": {
            "path": str(visual_path),
            "sha256": sha256_file(visual_path),
            "name": tokens["name"],
            "primary": color_hex(tokens["primary"]),
            "secondary": color_hex(tokens["secondary"]),
            "accent": color_hex(tokens["accent"]),
            "background": color_hex(tokens["background"]),
            "title_font": tokens["title_font"],
            "body_font": tokens["body_font"],
            "render_font": render_font,
        },
        "inputs": inputs,
        "geometry_source": "local-pillow",
    }


def main() -> int:
    args = build_parser().parse_args()
    try:
        result = run(args)
        exit_code = 0
    except Exception as exc:
        result = {
            "schema_version": 1,
            "tool": "ppt-gen-compose-template-overview",
            "generated_at": now(),
            "status": "failed",
            "errors": [str(exc)],
        }
        exit_code = 1
    print(json.dumps(result, ensure_ascii=False, indent=2), file=sys.stdout if exit_code == 0 else sys.stderr)
    if args.json_out:
        atomic_write_json(args.json_out, result)
    return exit_code


if __name__ == "__main__":
    raise SystemExit(main())
