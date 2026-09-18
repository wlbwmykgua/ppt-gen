"""Estimate spoken content duration; labels alone never prove speaking time."""
from __future__ import annotations

import argparse
import json
import math
import re
import unicodedata
from pathlib import Path


CJK = re.compile(r"[\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff\u3040-\u30ff\uac00-\ud7af]")
WORDS = re.compile(r"[^\W_]+(?:['’.-][^\W_]+)*", re.UNICODE)
SPOKEN_FIELDS = ("script", "spoken_script", "speaker_notes")


def parse_duration_seconds(value) -> int | None:
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return int(round(value)) if math.isfinite(value) and value >= 0 else None
    text = " ".join(unicodedata.normalize("NFKC", str(value or "")).split())
    if re.fullmatch(r"\d+(?:\.\d+)?", text):
        return int(round(float(text)))
    clock = re.search(r"(?<!\d)(\d{1,3}):(\d{2})(?!\d)", text)
    if clock:
        return int(clock.group(1)) * 60 + int(clock.group(2)) if int(clock.group(2)) < 60 else None
    minutes = re.search(r"(\d+(?:\.\d+)?)\s*分(?:钟)?(?:\s*(\d+)\s*秒)?", text)
    if minutes:
        return round(float(minutes.group(1)) * 60) + int(minutes.group(2) or 0)
    seconds = re.search(r"(\d+)\s*(?:秒|seconds?|sec\b)", text, re.IGNORECASE)
    if seconds:
        return int(seconds.group(1))
    minutes = re.search(r"(\d+(?:\.\d+)?)\s*(?:minutes?|mins?\b)", text, re.IGNORECASE)
    return round(float(minutes.group(1)) * 60) if minutes else None


def expected_entry_duration(entry: dict) -> int | None:
    for key in ("target_seconds", "duration_seconds", "expected_seconds", "duration", "expected_time", "time"):
        if key in entry:
            return parse_duration_seconds(entry[key])
    return None


def number(value, name: str, low: float, high: float) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(value) or not low <= value <= high:
        raise ValueError(f"{name} must be a finite number between {low} and {high}")
    return float(value)


def estimate_page(entry: dict, settings: dict) -> dict:
    script = next((entry[key] for key in SPOKEN_FIELDS if isinstance(entry.get(key), str) and entry[key].strip()), "")
    if not script:
        raise ValueError("a nonempty spoken script is required for timing")
    # Transition is spoken; goal, cue and source caution are presenter metadata.
    transition = entry.get("transition")
    if isinstance(transition, str) and transition.strip() and transition.strip() not in script:
        script += "\n" + transition
    cjk_rate = number(settings.get("cjk_chars_per_minute", 260), "cjk_chars_per_minute", 120, 500)
    word_rate = number(settings.get("words_per_minute", 150), "words_per_minute", 70, 240)
    pause_rate = number(settings.get("sentence_pause_seconds", 0.35), "sentence_pause_seconds", 0, 3)
    pauses = number(entry.get("pause_seconds", 0), "pause_seconds", 0, 3600)
    cjk_count = len(CJK.findall(script))
    word_count = len(WORDS.findall(CJK.sub(" ", script)))
    punctuation = len(re.findall(r"[。！？!?；;]|\.(?:\s|$)", script))
    seconds = cjk_count / cjk_rate * 60 + word_count / word_rate * 60 + punctuation * pause_rate + pauses
    return {"page": entry.get("page"), "cjk_characters": cjk_count, "words": word_count,
            "pause_seconds": round(punctuation * pause_rate + pauses, 1),
            "estimated_seconds": round(seconds, 1), "target_seconds": expected_entry_duration(entry)}


def validate_timing(data, entries, result: dict, requested_seconds=None, tolerance_percent: float = 10) -> None:
    if not isinstance(data, dict) or not isinstance(entries, list):
        return
    settings = data.get("timing", {})
    if not isinstance(settings, dict):
        result["errors"].append("speaker-notes timing must be an object")
        return
    try:
        if not entries or not all(isinstance(entry, dict) for entry in entries):
            raise ValueError("timing requires nonempty, structured slide entries")
        declared_request = settings.get("requested_seconds")
        if declared_request is not None:
            number(declared_request, "timing.requested_seconds", 1, 86400)
        if requested_seconds is not None:
            number(requested_seconds, "state requested_duration_seconds", 1, 86400)
            if declared_request is not None and declared_request != requested_seconds:
                raise ValueError("notes timing.requested_seconds conflicts with the user duration recorded in state")
        requested = requested_seconds if requested_seconds is not None else declared_request
        pages = [estimate_page(entry, settings) for entry in entries]
        for index, page in enumerate(pages, 1):
            target = number(page["target_seconds"], f"page {index} target_seconds", 1, 86400)
            estimate = page["estimated_seconds"]
            drift = abs(estimate - target) / target
            if estimate > target * 2 or estimate < target / 2:
                result["errors"].append(f"speaker_notes page {index}: script estimate {estimate}s is implausible for its {target:g}s target; revise script/time or document pauses")
            elif drift > 0.35:
                result["warnings"].append(f"speaker_notes page {index}: script estimate {estimate}s differs from {target:g}s target; rehearse or adjust pacing")
        total = sum(page["target_seconds"] for page in pages)
        estimate_total = round(sum(page["estimated_seconds"] for page in pages), 1)
        if requested is not None and abs(total - requested) / requested * 100 > tolerance_percent:
            result["errors"].append(f"speaker_notes target total {total}s differs from user-requested {requested}s by more than {tolerance_percent:g}%")
        if requested is not None and abs(estimate_total - requested) / requested > 0.35:
            result["warnings"].append(f"speaker_notes estimated spoken total {estimate_total}s is outside the requested {requested}s pacing range")
        result["metrics"]["speaker_timing"] = {"method": "text-rate-estimate-not-a-rehearsal",
            "requested_seconds": requested, "estimated_total_seconds": estimate_total,
            "target_total_seconds": total, "pages": pages}
    except (ValueError, TypeError) as exc:
        result["errors"].append(f"speaker-notes timing: {exc}")


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("ledger", type=Path)
    args = parser.parse_args()
    data = json.loads(args.ledger.read_text(encoding="utf-8"))
    result = {"errors": [], "warnings": [], "metrics": {}}
    validate_timing(data, data.get("slides"), result)
    print(json.dumps(result, ensure_ascii=False, indent=2))
    raise SystemExit(1 if result["errors"] else 0)


if __name__ == "__main__":
    main()
