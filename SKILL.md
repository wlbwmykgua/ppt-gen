---
name: ppt-gen
description: Create or continue an evidence-backed Word report, slide outline, image-first PPT, object-level editable PPT, and speaker notes. Use for any requested subset, existing reports or slide images, deck revisions, arbitrary start/stop/resume, or standalone deliverable QA. Chinese-first, exact 16:9, with offline OCR.
---

# Ppt Gen

Produce only requested outputs while keeping facts, page identity, and visual intent aligned. Chinese-first; otherwise follow the primary material. The established PPT route is intentionally image-first: approved slide images become the visual contract, then `image-to-editable-ppt` reconstructs editable objects. Do not silently replace it with a native-first authoring route.

## Load the exact companion skills

Read each applicable companion skill completely before its stage. Before production, use `grill-me` → `grilling` for the default idea interview, following [references/idea-interview.md](references/idea-interview.md). Resolve these skills in the current environment, not through machine-specific paths. Then route the production companions in this order and record the implementation in state:

1. Use `documents:documents` for DOCX creation/editing/render QA.
2. Use `pdf:pdf` for PDF extraction/render/inspection.
3. Load `pptx` whenever any PPTX is input or output; it owns PPTX inspection, package/render QA, and presentation-tool routing.
4. Use `presentations:Presentations` for native PPT packaging/editing only when selected under the `pptx` workflow; it does not replace the image-to-editable reconstruction stage.
5. Use built-in `imagegen` for style previews, six layout specimens, and final slide images.
6. Use `image-to-editable-ppt` only for reconstruction from approved slide images/image-based material.

Companion validation is additive. This orchestration layer may tighten but never waive it.

## Phase graph and deliverables

Treat phases as resumable targets, not a mandatory chain:

| ID | Purpose | Typical output |
|---|---|---|
| `brief` | route, constraints, privacy, authority | project brief/state |
| `report` | research and synthesis | substantive DOCX |
| `outline` | evidence into slide story | slide-copy ledger/outline |
| `style-options` | 2–3 visual directions | exact-16:9 previews |
| `template` | lock visual system | six layouts + deterministic overview |
| `image-deck` | produce approved rasters | slide images and/or image PPTX |
| `editable-deck` | reconstruct objects | editable PPTX |
| `deck-revision` | revise an existing editable deck | versioned revised PPTX |
| `speaker-notes` | final-order talk track | notes JSON/DOCX/PPT notes |
| `qa-package` | delivery acceptance | QA JSON/evidence |

Read [references/deliverable-contracts.md](references/deliverable-contracts.md) before execution, [references/data-contracts.md](references/data-contracts.md) before writing ledgers/handoffs, and [references/quality-gates.md](references/quality-gates.md) before accepting artifacts.

Record exact file `deliverables` separately from phase `targets`. A phase can yield one requested file without its siblings. Do not synthesize omitted upstream artifacts merely because the full graph contains them.

## Intake, source adaptation, and privacy

Follow [references/intake.md](references/intake.md) and, for a new or materially changed brief, [references/idea-interview.md](references/idea-interview.md). Default to a `grill-me`/`grilling` interview: resolve the idea in dependency-aware rounds, recommend answers, challenge material tradeoffs, and obtain confirmation of the agreed production brief before generating deliverables. Reuse supplied answers; do not cap the interview at three grouped questions or repeat it on every continuation.

An explicit “直接做 / 不要再问 / 全套自动完成，其他你定” delegates unresolved choices and bypasses nonessential interview/confirmation waits. “先问清楚，再自动做完” still requires the interview, then runs continuously after confirmation. Neither path bypasses privacy, authority, or QA gates. Intake is a pre-production gate for the requested route, not permission to add skipped upstream phases.

Before routing supplied files, run `scripts/source_adapter.py`. Inspect contents, preserve originals, hash authorities, and adopt `source-manifest.json`. Choose and record:

- `source_policy`: `faithful` or `verify-update`;
- `data_handling`: `standard` or `local-only`.

For a bare topic with no file, first persist the user's wording verbatim as a project-local `sources/user-brief.md` and adapt that file; do not fabricate a source manifest from memory. When web research adds authorities, save a small local evidence record for each source (title, URL, retrieval time, relevant locator/summary), hash it, and append it to the manifest so claim IDs never point to an unrecorded web result.

Decide `data_handling` before web research, image generation, OCR, or worker dispatch. `local-only` prohibits remote calls/uploads; do not promise identical visual fidelity if remote visual generation is unavailable.

Run `scripts/preflight.py` before expensive work. Check the actual required document/PPT/OCR/font tools and agent-only capabilities. For long runs, create/render a tiny Chinese DOCX/PPTX smoke artifact using the selected companions. Read [references/runtime.md](references/runtime.md) for runtime setup, offline preparation, caching, or sharing with another computer; preflight and execution share the same tool resolver.

## Initialize, resume, or retarget state

Use `scripts/project_state.py` for multi-phase, multi-page, or resumable work. Initialize with route, targets, deliverables, and mode; add authority inputs immediately. Example (confirm flags with `--help`):

```bash
python <skill-root>/scripts/project_state.py init <project-dir> \
  --name "<project>" --mode guided --source-policy verify-update \
  --data-handling standard --start-at report --stop-after qa-package \
  --deliverable report-docx --deliverable slide-images --deliverable editable-pptx \
  --deliverable speaker-notes-docx --deliverable qa-report
python <skill-root>/scripts/project_state.py add-input <project-dir> primary=/absolute/input/path
python <skill-root>/scripts/project_state.py show <project-dir>
```

State rules (use `--help` for the exact subcommand syntax):

- `current_stage` is derived from unresolved requested targets; never point it to a skipped phase.
- A phase becomes `completed` only with an existing artifact, acceptable QA, and any required per-page checkpoints.
- `completed` run status is derived; no pending/stale/failed target may remain.
- On “继续”, run `verify` first. Missing or hash-changed inputs/artifacts make the owning phase and transitive consumers stale.
- On stop, persist page checkpoints and `pause`; on continuation `resume` from the first valid unresolved target.
- When the user extends/reduces scope, use `retarget`; never discard accepted earlier work.
- When the user edits/replaces an artifact, preserve it and `adopt` it as the new authority or make a versioned output.
- For QA-only work, target `qa-package`, request only `qa-report`, and register each existing deliverable with `add-input --stage qa-package --artifact-role <validator-role>`. It remains an input, not a new production deliverable; see the example in `references/intake.md`.
- Record selected style, structured source/privacy/OCR policies, worker implementation, warnings, attempts, hashes, and validation evidence. Route implementation/attempt details through `record --preference worker_implementation=<name>` and `record --decision "attempt:<stage>:<summary>"`; route warnings and evidence through `set-stage --warning ... --artifact role=/absolute/path` (or page equivalents). Do not rely on chat memory. Keep state `policies` consistent with `source-manifest.json`, preflight, and `editable-handoff.json`.

After the intake gate, in `stepwise`, pause after every accepted requested phase. In `guided`, pause at a material choice, normally style selection. In `continuous`, advance automatically after gates; continuous means no nonessential waiting, not lower QA.

## Build source and copy authority

For bare-topic factual work, research enough for a substantive report under the chosen source policy. Prefer current primary/authoritative sources, retain locators/cutoff, and distinguish fact, user authority, and analysis in `claim-ledger.json`.

Once an outline exists, create `slide-copy-ledger.json`. It is the exact authority for titles, body text, values, units, charts, sources, and page identity. Never retype critical content from memory. Markdown versions may aid review but are derived views.

## Report and outline

Create a rich report rather than a stretched outline. Render the DOCX and inspect all pages. Convert the accepted report/authority into a page-by-page story with one judgment per slide, evidence, source, visual intent, and speaker takeaway. Use comparison + data + charts only when reliable evidence exists.

## Visual decision and deterministic template

Read [references/visual-system.md](references/visual-system.md).

- Guided/stepwise: generate 2–3 genuinely different exact-16:9 previews from identical representative copy, persist `style-options.json`, then wait for selection.
- Continuous: hard-gate, score, select, and record the reason automatically.
- Skip for faithful reconstruction, supplied fixed design, or a run ending before visuals.

After selection, write `visual-system.json`. Generate six separate exact-16:9 layout specimens and validate them. Then use `scripts/compose_template_overview.py` to create the 3×2 overview. Do not ask one image-generation call to guarantee six internal 16:9 thumbnails. When the user supplies or approves reuse of a complete existing template package, adopt its selected style and layouts with current hashes instead of regenerating them; a requested overview still contains all six layouts.

## Generate and register slide images

Use one built-in `imagegen` call per page, supplying the page ledger and selected visual system. Keep project-bound copies. Fully decode and normalize without distortion, reconcile all critical fields, inspect each page plus a natural-order contact sheet, and register accepted files in `slide-manifest.json`.

On revision/resume, use `scripts/page_reuse.py` to identify unchanged accepted rasters. Regenerate affected pages only; reuse is a candidate decision, never an automatic QA pass. Refresh state/page evidence and verify the final whole-deck order, copy, geometry, and visual consistency.

If `slide-images` is requested, deliver them. If `image-pptx` is requested, package the same manifest pages in exact manifest order as one full-canvas image per slide. Produce neither sibling unless requested.

## Build the editable deck with offline OCR

Before dispatch, write `editable-handoff.json`. It must include absolute paths to authority ledgers/source images, hashes, `ocr_policy: offline`, `ask_for_ocr_token: false`, `data_handling`, allowed image backend, worker implementation, and a prompt tail.

The user's offline selection overrides companion default intake: do not ask for or use a PaddleOCR token. Append the handoff path and policy tail to every page-worker prompt. Use ledger text to correct OCR. “Offline” describes OCR/text hints; if `standard`, the companion may use an allowed built-in visual backend for isolated assets. If `local-only`, disable remote image generation/uploads.

For prepare, use `scripts/offline_prepare.py` as documented in `references/runtime.md`. It forces `editppt prepare --no-text-hints`, then runs local page geometry hints and Apple Vision recognition, producing `offline-preparation.json`. Never use unguarded prepare or `editppt run hints`; clearing the token environment alone does not disable a configured token. Include the execution report and each page's `offline-recognition.json` path in the worker handoff.

Then follow `image-to-editable-ppt` per-page reconstruction/dispatch → record → finalize. Rebuild text, simple shapes, tables, and known-data charts as independent editable objects; reserve PNG assets for complex photography/texture/illustration. Require each page's `validation.json` and rendered source/editable comparison. Never accept a source-slide screenshot with empty, tiny, hidden, or unrelated overlays as editable.

## Revise an existing deck

For `deck-revision`, preserve the original and inspect real slide order, masters/layouts, native text, notes, relationships, and media. Record page/object requests in `revision-checkpoints.json`. Use the presentation tools selected by `pptx`, render every changed page and a whole-deck contact sheet, confirm requested edits and unchanged critical content, then save a distinct versioned `revised-pptx`. Final QA binds the original PPTX, revised PPTX, and revision checkpoints together; never validate the revised deck alone or overwrite the original.

## Write speaker notes

Derive `speaker-notes.json` from the final accepted reference deck order, not an old outline. Without a requested duration, choose a natural total (normally 20–40 seconds for cover, 45–90 seconds per content page). Include exact page/title, target time, goal, spoken script, cue, transition, and claim/source cautions. Render only requested notes formats. Final QA reconciles the notes ledger and rendered output against both `slide-copy-ledger.json` and the reference deck. If `ppt-notes` is requested, inspect the delivered PPTX's native embedded notes and slide-to-notes relationships; a matching JSON or DOCX alone is insufficient.

If the user specifies a duration, persist `requested_duration_seconds` with state `record --preference`. Use `scripts/speech_timing.py` to estimate the actual spoken text before rendering. Record language/rate assumptions and explicit demo pauses in the notes ledger; time labels must fit both the script and the user's target. Estimates are not a timed rehearsal.

## Validate and deliver

Run state `verify` immediately before accepting each deliverable phase, then run `scripts/validate_delivery.py` without `--structural-only`, passing `--project-state`, `source-manifest.json`, and `claim-ledger.json` plus that phase's complete hard-gate bundle. This permits an accepted report to feed a later deck; the final `qa-package` confirms that every requested deliverable still has current hash-bound QA. The verify checkpoint revision must equal the state revision at validation time. Editable output includes the handoff, slide/source manifests, copy ledger, and hash-bound source/editable/comparison renders; deck revision includes original, revised, object-level checkpoints, and hash-bound visual evidence; speaker notes include the notes ledger, slide-copy ledger, reference deck, and rendered DOCX evidence; report and image-PPTX delivery likewise pass their render-evidence JSON manifests defined in `references/data-contracts.md` (a naked PNG or render directory is not accepted). Requested `ppt-notes` includes native embedded-notes validation. An image-only reference deck also passes its approved `slide-manifest.json`; the validator binds the actual embedded page images and order as well as titles. `--structural-only` is diagnostic only and never counts as final-delivery QA, phase completion, or delivery authority. Perform the visual/open/render QA required by document/PPT skills. The validator must fail closed on malformed/empty packages, missing media/relationships, corrupt images, missing authorities/evidence, page/order/critical-copy mismatch, fake editability, OCR/privacy handoff mismatch, incomplete revision checkpoints, or notes mapping/embedding errors.

Editable final QA additionally passes `--offline-preparation`; the execution record must bind the current handoff, ordered source images, and local recognition outputs. Reuse of OCR/render cache never waives source hashes, semantic checks, final visual inspection, or companion QA.

Write `qa-report.json` with tri-state results and evidence. Deliver only requested artifacts, with absolute clickable paths, plus concise assumptions, selected style, source policy/cutoff, data handling, expected duration, QA status, and any warnings. Never call a warning-free Boolean pass when warnings remain.

## Common invocations

- “我只有一个想法，先用 grill-me 帮我问清楚，再开始做。”
- “先把方案问清楚，我确认后你就不停下全部做完。”
- “介绍一个主题，只做到 Word 报告。”
- “我有报告，从大纲开始；风格图让我选。”
- “只生成 PPT 图片，不要打包 PPTX。”
- “把这些图片离线转为可编辑 PPT，不要问 OCR token。”
- “修改这份现有 PPTX，只交付新版 PPT。”
- “根据最终 PPT 写逐页讲稿，时间你定。”
- “从主题开始全部做完，不要停，细节你决定。”
