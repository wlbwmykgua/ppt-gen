# Deliverable contracts

Read [data-contracts.md](data-contracts.md) for the JSON schemas that make these outputs resumable and testable. JSON ledgers are authoritative; DOCX, Markdown, images, and PPTX are rendered deliverables or human-readable views.

Every final non-`--structural-only` QA run supplies `source-manifest.json` and `claim-ledger.json` as top-level authorities in addition to the requested deliverable's complete evidence bundle. A nested pointer does not replace either authority.

## Deliverables are independent of phases

The user may request any subset. Record exact `deliverables` separately from execution `targets`:

- `report-docx`
- `outline`
- `style-options`
- `template-overview`
- `slide-images`
- `image-pptx`
- `editable-pptx`
- `revised-pptx`
- `speaker-notes-docx`
- `ppt-notes`
- `qa-report`

For example, `image-deck` may produce only `slide-images`, only `image-pptx`, or both. Never create an unrequested sibling output merely because it shares a phase.

## Report

Create a polished Word report whose structure fits the topic. Usually include an executive summary, background, key chronology or framework, evidence and comparisons, data or cases, interpretation, implications, conclusion, and source appendix. Distinguish fact from inference and preserve definitions, dates, units, denominators, source locators, and cutoff. Use `source-manifest.json` and `claim-ledger.json`; render and inspect every DOCX page.

## Slide outline

Give every page one central judgment. Record page number, exact title/body copy, story purpose, supporting facts/data, proposed visual, source IDs/footer, and speaker takeaway in `slide-copy-ledger.json`. Create a narrative arc rather than copying report headings.

## Style options

Create 2–3 separate exact-16:9 preview images from identical representative content so style—not content—drives comparison. Vary the whole visual system. Save every prompt, file hash, geometry, parameters, hard-gate result, score, selection, and reason in `style-options.json`.

## Template overview

Lock `visual-system.json` first. Build six real exact-16:9 layout pages—cover, section, narrative, comparison, data/chart, conclusion—and then compose them deterministically into one exact-16:9 3-column × 2-row overview with `scripts/compose_template_overview.py`. Show palette, type, grid, spacing, imagery, icon, chart, and footer rules. Do not trust a generated overview image to preserve six internal 16:9 rectangles.

## Slide images

- Use one consistent exact-16:9 canvas, normally 1920×1080 or 2048×1152.
- Name pages with zero-padded numeric prefixes such as `01-cover.png`.
- Keep page count/order identical to `slide-copy-ledger.json`.
- Reconcile every title, name, date, value, unit, and chart label before approval.
- Record paths, hashes, dimensions, critical fields, and approved full-bleed exceptions in `slide-manifest.json`.
- Preserve approved PNGs even when no image-based PPTX was requested.

## Image-based deck

Assemble the approved `slide-manifest.json` pages into a separate PPTX with one exact full-canvas image per slide, in manifest order. Validate package relationships, media existence, dimensions, and rendered page identity. Do not substitute the editable deck for this deliverable.

## Editable deck

The established route remains image-first: use `image-to-editable-ppt` on approved slide images. Before dispatch, write `editable-handoff.json` with offline OCR, privacy policy, authority-ledger paths, source image hashes, worker prompt tail, and implementation name. The user's offline choice wins; never ask for or use a PaddleOCR token in this workflow.

Use the guarded preparation wrapper from [runtime.md](runtime.md). Final QA includes
`--offline-preparation` with current hashes for the handoff, normalized source pages,
local geometry hints, and Apple Vision recognition. A policy written in a prompt
without this execution evidence is not sufficient.

Rebuild text as text boxes; simple lines/cards/arrows/icons/containers as native shapes; tables as native tables; and known-data charts as native charts. Use separate PNG assets only for complex photos, textures, and illustrations. Never deliver a full-page bitmap plus token editable overlays as “editable.” Require every page's companion `validation.json`, a rendered side-by-side comparison, and ledger reconciliation.

Offline describes OCR/text hints. If `data_handling` is `standard`, the companion may use an allowed built-in image backend for page-local assets. If it is `local-only`, disable remote image generation/uploads and report any resulting fidelity limitation.

## Deck revision

For an existing editable PPTX plus a revision request, use the `deck-revision` phase. Preserve the original; inspect actual slide order, master/layout use, notes, native text, media, and editability. Record requested changes as page/object checkpoints in `revision-checkpoints.json`. Produce a distinct, versioned `revised-pptx`; never overwrite or relabel the original as revised.

Final revision QA is a three-part contract: original PPTX + revised PPTX + revision checkpoints. Bind current hashes for both decks, require a unique `(page, object_id)` checkpoint and machine-checkable `expected_after`/`new_value` for every requested object edit, bind every evidence file by SHA-256, render every changed page plus a whole-deck contact sheet, and reconcile unchanged critical content, slide order, notes, relationships, and media against the original. Multiple object checkpoints on one page are valid. A structurally valid revised deck without the original or complete checkpoints is not an accepted `revised-pptx`.

## Speaker notes

Author `speaker-notes.json` from the final accepted reference deck order, then render requested formats. When the user asks only for “讲稿” and gives no format, default to an independent `speaker-notes-docx`; embed `ppt-notes` only when explicitly requested or clearly useful and supported. A notes DOCX contains a cover plus one section per slide with thumbnail, exact page/title, expected time, goal, natural script, cue, transition, and source caution. Never infer order from an obsolete outline.

Final notes QA must reconcile `speaker-notes.json` against both `slide-copy-ledger.json` and the accepted reference deck in the same run: page count, one-based order, normalized titles, duration totals, and claim/source cautions must agree. The notes ledger or DOCX cannot self-attest without those references.

Estimate spoken duration from the actual script and transition, language/rate settings,
and explicit demonstration pauses. Persist a user-requested duration in state as
`requested_duration_seconds`; no requested duration means the script determines a
natural total. A rate estimate is not a rehearsal, so disclose meaningful pacing
uncertainty rather than claiming an exact delivery time.

If `ppt-notes` is requested, validate the actual notes-bearing PPTX as a separate requested artifact. Inspect slide-to-notes relationships and native embedded text, then compare every slide's embedded notes with `speaker-notes.json`. A valid notes DOCX/JSON does not prove that notes were embedded; missing, empty, extra, stale, or wrong-slide notes block delivery.

## QA package

Create `qa-report.json` with tri-state overall and per-gate status, top-level source-manifest/claim-ledger hashes, requested-artifact hashes, evidence paths, errors, and warnings. QA runs for every requested deliverable even when `qa-package`/`qa-report` is not itself requested; in that case keep the record as internal acceptance evidence rather than adding it to the user-facing deliverables.

For QA-only imports, `qa-report` is the only new deliverable. Typed existing inputs
are registered at `qa-package` as described in `intake.md`; their normal evidence
bundles and current hashes remain mandatory. Completing QA does not mark the skipped
production stages complete or permit overwriting the supplied files.

Each deliverable phase is accepted by a non-`--structural-only` validation containing that phase's complete applicable hard-gate bundle from [quality-gates.md](quality-gates.md), including `--project-state` immediately after a no-drift `verify`. This lets a report be accepted before its downstream deck exists without weakening the later deck gate. Pass the dedicated report, image-PPTX, and speaker-notes render-evidence JSON flags when those formats are delivered. The final `qa-package` audits that every requested deliverable has a current accepted QA record whose artifact and authority hashes still match state; it does not demand an impossible all-artifacts-before-first-phase run. Include only evidence needed to reproduce each decision: authority contracts, revision or notes references when applicable, companion validations, hash-bound renders/contact sheets, and validator output. `passed_with_warnings` is usable only when no hard gate failed and every warning is disclosed; `failed` blocks delivery. `--structural-only` is diagnostic only and cannot mark a deliverable complete or authorize handoff.

## Suggested project layout

```text
<project>/
  ppt-gen-state.json
  source-manifest.json
  claim-ledger.json
  slide-copy-ledger.json
  style-options.json
  visual-system.json
  slide-manifest.json
  editable-handoff.json
  revision-checkpoints.json
  speaker-notes.json
  sources/
  01-report/
  02-outline/
  03-style-options/
  04-template/layouts/
  05-slide-images/
  06-image-deck/
  07-editable-deck/
  08-deck-revision/
  09-speaker-notes/
  qa/
```

Reuse a sensible existing project structure when present. Never move or overwrite original user files without permission.
