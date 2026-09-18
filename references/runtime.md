# Runtime, safe preparation, and repeatable checks

Read this reference when configuring a computer, preparing editable reconstruction,
using persistent caches, or testing changes to this skill. No command here installs
software or changes account configuration automatically.

## Runtime and distribution

Send the whole `ppt-gen` directory, including `scripts/`, `references/`, and `tests/`.
Do not bundle personal projects, caches, credentials, or proprietary companion skills.
The recipient needs the companion skills applicable to their selected stages; resolve
them in that environment rather than copying absolute paths from this Mac.

- Idea refinement uses `grill-me` → `grilling` when available. These are separately
  installed companions, not bundled copies; the scoped interview fallback in
  `references/idea-interview.md` keeps intake usable when sharing without them.
- Python 3.10+; `requirements.txt` lists the local helpers' Python dependency.
- Prefer the Codex bundled artifact Python returned by `load_workspace_dependencies`.
  DOCX/PPT authoring libraries belong to the selected companion runtime. Optional
  integration-test dependencies are in `tests/requirements.txt`.
- Final DOCX/PPTX comparison requires LibreOffice and Poppler (`pdftoppm`).
- Editable reconstruction requires the separately installed `image-to-editable-ppt`
  skill and `editppt` CLI with `prepare --no-text-hints` support.
- Apple Vision OCR requires macOS and Xcode Command Line Tools. The complete workflow
  has been tested on macOS only; do not advertise full Windows/Linux compatibility.
  Missing offline recognition blocks the dependent stage, never enables online OCR.

Preflight and rendering use `scripts/local_runtime.py`. Explicit executable overrides
are `PPT_GEN_LIBREOFFICE`, `PPT_GEN_PDFTOPPM`, `PPT_GEN_SWIFTC`, and `PPT_GEN_EDITPPT`.
Set an override only to a verified executable; an invalid override fails closed.
Record preflight JSON with the exact interpreter and resolved paths before production.
Preflight does not prove the agent can call imagegen or dispatch workers.

Older accepted QA records remain readable under their historical contract. Current
validation writes `validation_contract_version: 2`. Historical editable acceptance
without guarded-preparation evidence is disclosed as legacy, never relabeled as v2;
newly generated or revalidated editable outputs must supply the current execution
record. Do not hand-edit old QA JSON to claim the new contract.

## Guarded offline preparation

After writing the current `editable-handoff.json`, use a fresh versioned run directory:

```bash
python <skill-root>/scripts/offline_prepare.py \
  --handoff /absolute/project/editable-handoff.json \
  --job-dir /absolute/project/editable-run-v1 \
  --report /absolute/project/qa/offline-preparation-v1.json \
  --cache-dir /absolute/project/.ppt-gen-cache
```

`--dry-run` checks input hashes and prints the guarded command without running it.
The wrapper passes `--no-text-hints` regardless of existing environment/config tokens,
then explicitly calls `editppt page hints` (local geometry) and Apple Vision (recognized
text, confidence, normalized bounding boxes). It never calls `editppt run hints` or a
remote OCR service. The execution report records current source, hint, recognition,
and handoff hashes. Register it in state and pass it as `--offline-preparation` to
final validation. Give workers their page's `offline-recognition.json` alongside the
companion's geometry hints and the authoritative copy ledger.

Existing run/report files are preserved. A failed prepare must be diagnosed before
choosing a new versioned run; never overwrite an active reconstruction run.

The currently reviewed `editppt` version has no truly disabled image-backend contract.
For `data_handling: local-only`, the wrapper therefore stops before prepare. Use an
explicitly approved local-only reconstruction implementation that can record a disabled
backend; do not configure a remote-capable fallback and label it disabled. This is
separate from normal `standard` processing, whose OCR remains offline while built-in
image generation may be used for isolated assets.

## OCR/render caching

Without `--cache-dir`, the helpers share a process-local cache that is cleaned on
normal exit. A project-private `--cache-dir` enables reuse across continuations.
Cache keys include source bytes, backend/compiler identity, OCR implementation,
render DPI, font inventory, and relevant locale settings. Cached output hashes are
rechecked; missing or corrupted entries rebuild. Concurrent cache writes are locked.

Keep caches outside directories registered as source authorities. Cache files can
contain source text/images: keep them local, do not share them with the skill, and
do not mistake them for delivered artifacts. Font installation, tool updates, source
changes, and image changes invalidate their corresponding keys. The cache only avoids
repeating deterministic work; it never supplies an acceptance decision.

## Regression tests

From any directory, use an interpreter with the appropriate dependencies:

```bash
python -B -m unittest discover -s <skill-root>/tests -v
PPT_GEN_RUN_INTEGRATION=1 python -B -m unittest discover -s <skill-root>/tests -v
```

The default suite uses isolated local fixtures and no network or image-generation
calls. Integration tests additionally exercise actual Word/PPTX rendering, guarded
editppt preparation, and Apple Vision OCR on synthetic content. They require the
tools above plus `tests/requirements.txt`; they still perform no remote generation.
Use `PPT_GEN_TEST_ARTIFACTS_DIR` only for an explicit scratch directory when retaining
smoke previews for visual inspection. Do not count skipped integrations as passed.
