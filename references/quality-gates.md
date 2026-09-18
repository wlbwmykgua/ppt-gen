# Quality gates

Use `passed`, `passed_with_warnings`, or `failed`. A hard failure blocks every dependent phase in every run mode. Warnings are never silently converted to pass: a phase's state QA value must match its accepted `qa-report.json`, and warnings from accepted earlier phase reports remain visible in later final validation and the aggregate QA package.

## Cross-workflow authority

- Validate machine contracts from [data-contracts.md](data-contracts.md) before rendered outputs.
- Every final non-`--structural-only` run receives `source-manifest.json` and `claim-ledger.json` as top-level validator inputs. A nested reference does not satisfy this gate.
- Re-hash inputs/artifacts on resume and delivery; missing/changed authority invalidates transitive consumers.
- Verify unstable claims under `verify-update`; preserve supplied content under `faithful` while disclosing material conflicts.
- Preserve every value's unit, period, denominator, definition, source locator, and cutoff.
- Never invent chart data. Use timeline, matrix, flow, or qualitative structure when evidence is unavailable.
- Mark inference as analysis and unverified supplied statements as user authority.
- Treat image-generated and OCR text/data as untrusted until ledger reconciliation.
- Page count, order, titles, names, dates, values, units, and `critical_fields`: 100% agreement.

## Final-delivery hard-gate bundles

Final QA validates the deliverable(s) built at the current phase together with all of their authorities and evidence in one non-`--structural-only` run. Universal requirements are a passing state `verify`, current top-level `source-manifest.json` and `claim-ledger.json`, the current artifact(s), current hashes, and evidence paths that exist. In a multi-phase run, each requested deliverable receives its own phase acceptance before downstream work proceeds; the final `qa-package` audits that every requested deliverable still has current, hash-bound acceptance evidence. The following additions are mandatory when applicable:

| Requested deliverable | Same-run hard-gate bundle |
|---|---|
| `report-docx` | report DOCX; source manifest; claim ledger with report-bound evidence; project state/fresh verify; `--report-render-evidence` |
| `outline` | outline artifact; slide-copy ledger; source manifest; claim ledger |
| `style-options` | style-options ledger; 2–3 hashed exact-16:9 candidates; selection when advancing |
| `template-overview` | selected visual system; six exact-16:9 layouts; deterministic composer evidence; overview |
| `slide-images` | slide images; slide manifest; slide-copy ledger; page-level text/visual evidence |
| `image-pptx` | image PPTX; slide manifest/source images; slide-copy ledger; package/order; `--image-pptx-render-evidence` |
| `editable-pptx` | editable PPTX; slide-copy ledger; slide manifest/source images; editable handoff; guarded offline-preparation execution record; every page and final companion validation; rendered comparisons |
| `revised-pptx` | original PPTX; revised PPTX; revision checkpoints; changed-page renders; whole-deck contact sheet |
| `speaker-notes-docx` | notes DOCX; speaker-notes ledger; slide-copy ledger; accepted reference deck; `--speaker-notes-render-evidence` |
| `ppt-notes` | notes-bearing PPTX; speaker-notes ledger; slide-copy ledger; accepted reference deck; native embedded-notes inspection |

Omitting any applicable bundle member for the deliverable(s) under review is a hard failure. Do not combine unrelated or stale Boolean results. A later phase may retain an earlier deliverable's QA only while the state still binds its exact artifact/authority hashes and transitive invalidation has not marked it stale.

## Phase gates

| Phase | Deterministic checks | Visual/semantic checks | Hard failures |
|---|---|---|---|
| `brief` / inputs | source manifest, hashes, route/mode/targets/deliverables/privacy recorded; preflight completed | audience/purpose sufficient | corrupt/missing authority, prohibited remote handling |
| `report` | valid nonempty DOCX; XML/relationships/media valid; no unresolved placeholders; headings/sources | render and inspect every page, tables, fonts, overflow, blanks | malformed/empty file, unsupported critical claim, visible defect |
| `outline` | contiguous pages; slide ledger schema; source mapping; one judgment/page | coherent story and suitable visual logic | unsupported conclusion or chart without real data |
| `style-options` | 2–3 fully decoded exact-16:9 files; hashes/prompts/parameters/scores; selection when advancing | same comparison content, readable Chinese, contrast, extensibility | wrong ratio, false critical copy, unusable contrast, no selected style at gate |
| `template` | visual-system schema; six exact-16:9 source layouts; deterministic 3×2 overview | all page types coherent and useful | missing layout, non-16:9 source, internal geometry guessed by one generated board |
| `image-deck` | natural contiguous order; full decode; slide manifest/hash/dimensions; image PPTX package/order/media when requested | contact sheet plus every page full-size | missing page, corrupt image, wrong ratio/fact/order, clipping/unreadability |
| `editable-deck` | valid nonempty PPTX; true slide order; package rels/media; exact ratio/count; ledger text; every page companion validation | render beside source; layers, coordinates, fonts, z-order, selectable objects | full-slide source bitmap with token overlays, hidden/empty editability, missing objects/media, text/shape overflow |
| `deck-revision` | distinct original/revised paths and hashes; revision-checkpoints schema; every requested page/object checkpoint passed; package/order/media/notes valid | render every changed page and whole-deck contact sheet; inspect unchanged critical content | original overwritten; missing/failed/unbound checkpoint; requested change absent; unrelated critical content changed |
| `speaker-notes` | speaker-notes schema; one contiguous section per slide; exact mapping to slide-copy ledger and accepted reference deck; independent physical DOCX render-page count; computed total duration; embedded-notes package checks when `ppt-notes` is requested | natural speech, cues, transitions, no new claims; render DOCX and inspect native PPT notes when applicable | ledger/deck/page/title mismatch; unverified claim; requested duration off by >10%; missing, empty, extra, or misattached embedded notes |
| `qa-package` | validator tri-state; artifact hashes; evidence paths exist; no stale targets | warnings accurately disclosed | any unresolved hard gate or stale/missing requested deliverable |

## Editable-deck evidence

- `editable-handoff.json` records offline OCR/no-token and privacy choices. The separate
  `offline-preparation.json` proves the guarded prepare command and binds actual local
  recognition outputs to current inputs; a prompt-only declaration is insufficient.
- Final-delivery validation supplies that handoff, the slide-copy ledger, slide manifest/source images, and companion page validations together. Structural-only inspection is diagnostic and cannot satisfy this gate.
- Each page has `image-to-editable-ppt` record/finalize `validation.json` with `passed: true` and matching source/output identity.
- Native text matches ledger copy and all critical fields.
- Simple shapes/tables and known-data charts are native objects where required.
- A source-slide image covering about the whole canvas fails even if accompanied by empty, tiny, transparent, off-canvas, or unrelated objects. A second picture is not proof of editability.
- Complex photos/textures may be raster assets; an explicitly approved full-bleed photo is not the source-slide screenshot and must be declared in `slide-manifest.json`.
- Rendered source/editable pairs are inspected for visual fidelity in addition to structure.

## Speaking-time checks

The `speech_timing.py` helper estimates CJK characters, other-language words, spoken
transitions, sentence pauses, and explicit `pause_seconds`. Goal/cue/source-caution
metadata is not counted as speech. Page/total duration labels must still reconcile.
A user-requested total in state must match the planned total within 10% by default.
A page whose script estimate is above twice or below half its target fails as
implausible; intermediate drift above 35% is a rehearsal/pacing warning. These are
workflow heuristics, not a claim of measured performance. Do not manipulate rate or
pause settings merely to obtain a pass.

## Package and parser strictness

- Fully decode PNG/JPEG; header-only or truncated files fail.
- Run ZIP CRC checks; malformed XML and missing internal relationship targets fail.
- Read PPT slide order from `presentation.xml` relationships, not filenames.
- Reject zero-slide PPTX and empty DOCX.
- Reconstruct paragraph text across split XML runs for ledger/placeholder checks.
- Detect unresolved template tokens such as `{{...}}`, `${...}`, `<placeholder>`, or standalone `TODO/TBD`; ordinary prose containing the word “todo” is not automatically a placeholder.

## Structural-only diagnostics

`--structural-only` exists only to diagnose parser, ZIP/XML, relationship, slide-count, or basic object-structure problems. It may waive authority-ledger, semantic reconciliation, reconstruction-evidence, revision-checkpoint, or notes-binding requirements only for diagnosis. It never counts as final-delivery QA, even if the command exits zero or reports `passed`/`passed_with_warnings`.

A structural-only result must be labeled as diagnostic, retain a warning that final gates were not run, and must not:

- mark a stage or run `completed`;
- set final QA to `passed` or `passed_with_warnings` in project state;
- satisfy a requested deliverable's acceptance gate;
- authorize a user-facing handoff.

Rerun without `--structural-only` and supply the complete applicable bundle before delivery.

## Recovery and delivery

- Repair/regenerate only failed pages, preserving accepted checkpoints.
- After the same root cause fails twice, change the method before retrying.
- `passed_with_warnings` may continue only when correctness/usability remains intact; record affected artifact/page and mitigation.
- Never bypass factual errors, unreadable critical text, wrong ratio/order, fake editability, missing authority, credentials, paid actions, or privacy constraints.
- Before accepting each deliverable phase, run state `verify`, perform one non-structural validation with top-level source/claim authorities and that phase's complete applicable bundle, open/render its requested artifacts with the responsible companion skill, and write a phase QA record. Before final handoff, the `qa-package` confirms all requested deliverable records and hashes are still current.

Final delivery reports requested files only, selected style, inferred assumptions, source policy/cutoff, data-handling choice, expected speaking time, tri-state QA, and remaining warnings.
