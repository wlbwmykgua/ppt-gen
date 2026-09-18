"""Shared local tool discovery, hash-checked OCR/render caches, and cleanup.

No network or credential configuration is used here. Persistent caching is
opt-in; the default cache is process-local and removed on normal exit.
"""
from __future__ import annotations

import atexit
import hashlib
import json
import os
import platform
import shutil
import subprocess
import tempfile
from pathlib import Path
from typing import Callable

from file_lock import exclusive_lock


TOOL_NAMES = {
    "libreoffice": ("soffice", "libreoffice"),
    "pdftoppm": ("pdftoppm",),
    "swiftc": ("swiftc",),
    "editppt": ("editppt",),
}
TOOL_PATHS = {
    "libreoffice": (
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        "/Applications/LibreOffice.app/Contents/MacOS/libreoffice",
        "/opt/homebrew/bin/soffice", "/usr/local/bin/soffice",
    ),
    "pdftoppm": ("/opt/homebrew/bin/pdftoppm", "/usr/local/bin/pdftoppm"),
    "swiftc": ("/usr/bin/swiftc",),
    "editppt": (),
}
_TEMP_CACHE: tempfile.TemporaryDirectory | None = None
_CACHE: LocalCache | None = None


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def locate_tool(name: str) -> str | None:
    """An explicit task override wins; an invalid override fails closed."""
    override = os.environ.get("PPT_GEN_" + name.upper())
    if override:
        path = Path(override).expanduser()
        return str(path.resolve()) if path.is_file() and os.access(path, os.X_OK) else None
    for command in TOOL_NAMES[name]:
        found = shutil.which(command)
        if found:
            return str(Path(found).resolve())
    for raw in TOOL_PATHS[name]:
        path = Path(raw)
        if path.is_file() and os.access(path, os.X_OK):
            return str(path.resolve())
    return None


def font_roots() -> list[Path]:
    roots = [Path("/System/Library/Fonts"), Path("/Library/Fonts"),
            Path.home() / "Library/Fonts", Path("/usr/share/fonts"),
            Path("/usr/local/share/fonts"), Path.home() / ".local/share/fonts",
            Path.home() / ".fonts"]
    roots.extend(sorted(Path("/System/Library/AssetsV2").glob("com_apple_MobileAsset_Font*")))
    return roots


def font_fingerprint() -> str:
    records = []
    for root in font_roots():
        if not root.is_dir():
            continue
        for current, dirs, names in os.walk(root, followlinks=False):
            dirs.sort()
            for name in sorted(names):
                if Path(name).suffix.lower() not in {".ttf", ".otf", ".ttc", ".dfont", ".conf"}:
                    continue
                path = Path(current) / name
                try:
                    stat = path.stat()
                    records.append((str(path), stat.st_size, stat.st_mtime_ns, stat.st_ctime_ns))
                except OSError:
                    records.append((str(path), "unreadable"))
    return hashlib.sha256(json.dumps(records).encode()).hexdigest()


def tool_identity(path: str, version_args: tuple[str, ...]) -> dict:
    completed = subprocess.run([path, *version_args], capture_output=True, text=True,
                               timeout=15, check=False)
    return {"path": path, "sha256": sha256(Path(path)),
            "version": (completed.stdout + completed.stderr).strip(),
            "platform": platform.platform()}


class LocalCache:
    def __init__(self, root: Path):
        self.root = root.expanduser().resolve()
        self.root.mkdir(parents=True, exist_ok=True, mode=0o700)

    def _read(self, entry: Path, key: str) -> list[Path] | None:
        if entry.is_symlink():
            return None
        try:
            data = json.loads((entry / "cache.json").read_text(encoding="utf-8"))
            if data.get("key") != key or not data.get("files"):
                return None
            paths = []
            for record in data["files"]:
                name = record["name"]
                if not isinstance(name, str) or Path(name).name != name or name in {".", ".."}:
                    return None
                path = entry / name
                if path.is_symlink() or not path.is_file() or sha256(path) != record["sha256"]:
                    return None
                paths.append(path)
            return paths
        except (OSError, ValueError, KeyError, TypeError, AttributeError):
            return None

    def get(self, kind: str, identity: dict, build: Callable[[Path], list[Path]]) -> list[Path]:
        key = hashlib.sha256(json.dumps({"schema": 1, "kind": kind, **identity},
                                      sort_keys=True).encode()).hexdigest()
        with exclusive_lock(self.root / (key + ".lock")):
            return self._get_locked(key, build)

    def _get_locked(self, key: str, build: Callable[[Path], list[Path]]) -> list[Path]:
        entry = self.root / key
        cached = self._read(entry, key)
        if cached is not None:
            return cached
        with tempfile.TemporaryDirectory(prefix=".build-", dir=self.root) as staging:
            stage = Path(staging)
            outputs = build(stage)
            if not outputs or any(p.parent != stage or p.is_symlink() or not p.is_file() for p in outputs):
                raise RuntimeError("cache builder produced no files or nonlocal output")
            data = {"key": key, "files": [{"name": p.name, "sha256": sha256(p)} for p in outputs]}
            (stage / "cache.json").write_text(json.dumps(data), encoding="utf-8")
            if entry.exists():
                # Only this exact digest-addressed cache entry is disposable.
                if entry.is_symlink() or not entry.is_dir():
                    raise RuntimeError("refusing to replace a non-directory cache entry")
                cached = self._read(entry, key)
                if cached is not None:
                    return cached
                shutil.rmtree(entry)
            try:
                stage.rename(entry)
            except OSError:
                cached = self._read(entry, key)
                if cached is not None:
                    return cached
                raise
        cached = self._read(entry, key)
        if cached is None:
            raise RuntimeError("cache output failed its hash check")
        return cached


def configure_cache(directory: Path | None = None) -> LocalCache:
    global _TEMP_CACHE, _CACHE
    if directory is None:
        if _TEMP_CACHE is None:
            _TEMP_CACHE = tempfile.TemporaryDirectory(prefix="ppt-gen-cache-")
            atexit.register(_TEMP_CACHE.cleanup)
        directory = Path(_TEMP_CACHE.name)
    _CACHE = LocalCache(directory)
    return _CACHE


def cache() -> LocalCache:
    return _CACHE if _CACHE is not None else configure_cache()


def recognize_image(path: Path) -> dict:
    if platform.system() != "Darwin":
        raise RuntimeError("Apple Vision offline OCR requires macOS; no online fallback is allowed")
    compiler = locate_tool("swiftc")
    script = Path(__file__).with_name("offline_ocr.swift")
    if not compiler or not script.is_file():
        raise RuntimeError("offline OCR requires swiftc and offline_ocr.swift")
    identity = {"script": sha256(script), "compiler": tool_identity(compiler, ("--version",))}

    def compile_ocr(stage: Path) -> list[Path]:
        executable = stage / "offline-ocr"
        subprocess.run([compiler, str(script), "-O", "-o", str(executable)],
                       capture_output=True, text=True, timeout=120, check=True)
        return [executable]

    executable = cache().get("ocr-executable", identity, compile_ocr)[0]

    def run_ocr(stage: Path) -> list[Path]:
        completed = subprocess.run([str(executable), str(path)], capture_output=True,
                                   text=True, timeout=120, check=True)
        payload = json.loads(completed.stdout)
        if payload.get("backend") != "apple-vision-offline" or not isinstance(payload.get("lines"), list):
            raise RuntimeError("unexpected offline OCR result")
        output = stage / "ocr.json"
        output.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
        return [output]

    output = cache().get("ocr-image", {**identity, "image": sha256(path)}, run_ocr)[0]
    return json.loads(output.read_text(encoding="utf-8"))


def render_pages(artifact: Path) -> list[Path] | None:
    office, poppler = locate_tool("libreoffice"), locate_tool("pdftoppm")
    if not office or not poppler or artifact.suffix.lower() not in {".docx", ".pptx"}:
        return None
    identity = {"artifact": sha256(artifact), "suffix": artifact.suffix.lower(), "dpi": 150,
                "office": tool_identity(office, ("--version",)),
                "poppler": tool_identity(poppler, ("-v",)), "fonts": font_fingerprint(),
                "locale": {name: os.environ.get(name) for name in ("LANG", "LC_ALL", "LC_CTYPE", "TZ", "FONTCONFIG_FILE", "FONTCONFIG_PATH")}}

    def render(stage: Path) -> list[Path]:
        with tempfile.TemporaryDirectory(prefix="ppt-gen-office-") as work:
            workdir = Path(work)
            subprocess.run([office, "--headless", f"-env:UserInstallation={(workdir / 'profile').as_uri()}",
                            "--convert-to", "pdf", "--outdir", str(workdir), str(artifact)],
                           capture_output=True, text=True, timeout=120, check=True)
            pdfs = list(workdir.glob("*.pdf"))
            if len(pdfs) != 1:
                raise RuntimeError("office render produced no unique PDF")
            subprocess.run([poppler, "-png", "-r", "150", str(pdfs[0]), str(stage / "page")],
                           capture_output=True, text=True, timeout=120, check=True)
        return sorted(stage.glob("page-*.png"), key=lambda p: int(p.stem.rsplit("-", 1)[1]))

    return cache().get("office-render", identity, render)
