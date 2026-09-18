# Machine-readable contracts

JSON files are the execution authority. Markdown tables and prose are optional human-readable views derived from them; they never override JSON. Use absolute paths for files that workers must open, UTF-8, stable IDs, one-based page numbers, ISO-8601 timestamps, and SHA-256 hashes. Preserve unknown fields when updating a contract.

## Contents

- Source and evidence authority
- Slide and visual authority
- Editable reconstruction handoff
- Deck revision authority
- Speaker notes and QA
- Final-delivery validation bundle
- Validation rules

## Source and evidence authority

`source_adapter.py` may seed the three authority ledgers, but seeded ledgers
carry `status: "draft"`. They establish page order and source identity only.
Final validation rejects them until the agent replaces draft claim reasons with
verified claim mappings, records the real source cutoff, completes copy/notes
fields, and removes the draft status.

`source-manifest.json` records what entered the workflow and how it may be used:

```json
{
  "schema_version": 1,
  "source_policy": "faithful",
  "data_handling": "standard",
  "ocr_policy": "offline",
  "external_ocr_allowed": false,
  "ask_for_ocr_token": false,
  "primary_source_id": "src-001",
  "inferred_start": "outline",
  "sources": [{
    "id": "src-001", "role": "primary", "path": "/abs/report.docx",
    "sha256": "...", "media_type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    "authority_rank": 1, "adapter": "docx", "suggested_start": "outline",
    "properties": {"pages": null, "text_characters": 8400}
  }],
  "warnings": []
}
```

`source_policy` is `faithful` (do not silently update the supplied story) or `verify-update` (verify unstable claims and update with cited evidence). `data_handling` is `standard` or `local-only`; decide it before any remote call.

Every `sources[].path` is an existing absolute local file or directory whose hash is rechecked. For a bare prompt, persist the exact user brief locally. For a web authority, persist a compact evidence record containing its title, canonical URL, retrieval time, locator, and relevant summary, then list and hash that record; do not use an ephemeral search-result identifier as source authority.

`claim-ledger.json` records every consequential factual or analytical statement:

```json
{
  "schema_version": 1,
  "source_cutoff": "2026-08-12",
  "claims": [{
    "id": "claim-001", "statement": "Exact approved wording",
    "kind": "fact", "value": 42, "unit": "%", "period": "2025",
    "denominator": "survey respondents", "source_ids": ["src-001"],
    "source_locator": "p. 8", "verification": "verified",
    "used_in_report": true, "report_evidence": ["Exact approved wording"],
    "used_on_slides": [4],
    "slide_evidence": {"4": ["Exact approved wording"]},
    "notes_evidence": {"4": ["Exact approved wording"]}
  }]
}
```

Allowed `kind`: `fact`, `analysis`, `user-provided`. Allowed `verification`: `verified`, `user-authority`, `pending`, `unsupported`. `unsupported` blocks downstream use; `pending` blocks high-risk or time-sensitive claims. For report delivery, every claim must declare `used_in_report`; at least one resolved claim must be true and its `report_evidence` (or exact statement) must be found in the report text. For every slide use, `used_on_slides`, the page's `claim_ids`, and exact `slide_evidence` agree in both directions. When the approved slide wording paraphrases the claim statement, provide page-keyed `slide_evidence`; use `notes_evidence` when the spoken wording differs again. The validator must find those snippets in the corresponding authoritative copy/notes.

For final-delivery QA, `source-manifest.json` and `claim-ledger.json` are first-class top-level validator inputs. Pass both directly as `--source-manifest` and `--claim-ledger`; a path nested inside `editable-handoff.json`, a revision checkpoint, a notes ledger, or another artifact does not substitute for either top-level authority. Re-hash both files, require every `claim-ledger.json.claims[].source_ids` value to resolve to `source-manifest.json.sources[].id`, and reject any delivered claim whose verification state is not allowed for that use.

## Slide and visual authority

`slide-copy-ledger.json` is the exact text/data authority after the outline exists:

```json
{
  "schema_version": 1,
  "deck": {"title": "...", "canvas": "16:9", "source_cutoff": "2026-08-12"},
  "slides": [{
    "page": 1, "id": "slide-01", "title": "Exact title",
    "purpose": "One central judgment", "body_copy": ["Exact body line"],
    "visual_type": "comparison", "critical_fields": ["2026", "42%"],
    "data": null, "source_ids": ["src-001"], "footer": "Source: ...",
    "claim_ids": ["claim-001"], "speaker_takeaway": "One sentence"
  }]
}
```

Every slide records an explicit `claim_ids` array. Claims and
`claims[].used_on_slides` must agree in both directions. A cover, section, or
other genuinely non-claim page may use an empty array only with a nonempty
`claim_not_applicable_reason`; omission is not equivalent to “none.” A
speaker-notes page may reference only claim IDs approved on the corresponding
slide and follows the same explicit-empty rule.

`slide-manifest.json.text_verification_evidence` is an independent JSON review
record, never the slide image itself. It uses `schema_version: 1`,
`tool: "ppt-gen.slide-text-review"` (or `offline-ocr-review`), `passed: true`,
the exact `image_path` and `image_sha256`, an `observed_text` array, and
`reviewer: {"method": "manual-visual"}` or an offline-OCR method with
`ocr_policy: "offline"`. This record supplies review provenance, but it cannot
self-authenticate the pixels: final validation independently runs the bundled
local Apple Vision OCR verifier on the current image and reconciles its output
to the title, body, footer, and critical fields. If the offline verifier is
unavailable or does not see the authoritative copy, delivery fails closed.
Merely repeating words inside the manifest or review JSON is not evidence.

`style-options.json` makes the style gate resumable:

```json
{
  "schema_version": 1,
  "comparison_content_id": "slide-03",
  "generation_provenance": {"data_handling": "local-only", "backend": "local-native", "remote_generation": false, "uploaded": false},
  "candidates": [{
    "id": "A", "image": "/abs/A.png", "sha256": "...",
    "width": 1920, "height": 1080, "prompt_file": "/abs/A-prompt.txt",
    "parameters": {"palette": ["#111827", "#22D3EE"], "type": "editorial"},
    "hard_gate": "passed", "score": 88
  }],
  "selected_id": null, "selected_reason": null
}
```

`selected_id: null` is a valid accepted stopping point when the requested run
ends at style previews. It produces a disclosed warning and leaves downstream
visual work paused. Any route that includes template/image/editable targets
must populate both `selected_id` and `selected_reason` before advancing.

When `data_handling` is `local-only`, every visual authority created or adopted
at style, visual-system, and slide-manifest stages adds
`generation_provenance` with `data_handling: "local-only"`, one of
`backend: "local-native" | "local-pillow" | "user-provided" | "disabled"`,
and both `remote_generation: false` and `uploaded: false`. Missing provenance
fails closed; this records that no remote image service or upload was used.

`visual-system.json` is selected before the template overview and full deck:

```json
{
  "schema_version": 1,
  "name": "Selected direction",
  "source_style_options": "/abs/style-options.json",
  "source_style_options_sha256": "...",
  "selected_id": "A",
  "generation_provenance": {"data_handling": "local-only", "backend": "local-native", "remote_generation": false, "uploaded": false},
  "palette": {"background": "#F7F7F4", "primary": "#111827", "secondary": "#2563EB", "accent": "#F59E0B"},
  "typography": {"title_font": "Source Han Sans SC", "body_font": "Source Han Sans SC", "title_weight": 700, "body_weight": 400},
  "geometry": {"canvas": [1920, 1080], "margin": 96, "grid_columns": 12, "gutter": 24},
  "imagery": {"treatment": "high-contrast crop"},
  "charts": {"style": "direct labels, no 3D"},
  "icons": {"style": "simple outline"},
  "footer": {"position": "bottom-left", "minimum_size_pt": 10},
  "layouts": ["cover", "section", "narrative", "comparison", "data-chart", "conclusion"]
}
```

`slide-manifest.json` binds approved rasters to page identity and exact copy:

```json
{
  "schema_version": 1,
  "generation_provenance": {"data_handling": "local-only", "backend": "local-native", "remote_generation": false, "uploaded": false},
  "slides": [{
    "page": 1, "id": "slide-01", "path": "/abs/01-cover.png",
    "sha256": "...", "width": 1920, "height": 1080,
    "title": "Exact title", "critical_fields": ["2026", "42%"],
    "observed_text": ["Exact title", "Exact body line", "2026", "42%"],
    "text_verified": true, "text_verification_evidence": "/abs/qa/page-01-review.json",
    "allow_full_bleed_image": false
  }]
}
```

## Editable reconstruction handoff

Write `editable-handoff.json` before invoking `image-to-editable-ppt` and include its absolute path and policy tail in every page-worker prompt:

```json
{
  "schema_version": 1,
  "ocr_policy": "offline",
  "ask_for_ocr_token": false,
  "source_policy": "faithful",
  "data_handling": "standard",
  "image_backend": "builtin-allowed",
  "source_manifest": "/abs/source-manifest.json",
  "slide_copy_ledger": "/abs/slide-copy-ledger.json",
  "slide_manifest": "/abs/slide-manifest.json",
  "hashes": {
    "source_manifest": "...",
    "slide_copy_ledger": "...",
    "slide_manifest": "..."
  },
  "source_images": [{"path": "/abs/01-cover.png", "sha256": "..."}],
  "worker_prompt_tail": "Use offline OCR hints. Never request or use a PaddleOCR token. Correct all text and data from the listed ledgers.",
  "implementation": "image-to-editable-ppt"
}
```

If `data_handling` is `local-only`, set `image_backend` to `disabled` and forbid uploads. The workflow-level user choice overrides a companion's default intake: do not ask again for an OCR token when `ask_for_ocr_token` is false.

Run `scripts/offline_prepare.py` using this handoff; see `runtime.md` for invocation
and the current dependency's local-only limitation. Its `offline-preparation.json`
uses `schema_version: 1`, `tool: "ppt-gen.offline-prepare"`, `passed: true`, a current
`handoff: {path, sha256}`, the same ordered `source_images`, a `prepare_command`
containing `--no-text-hints`, `ocr_policy: "offline"`, `external_ocr_used: false`,
`backend: "apple-vision-offline"`, and contiguous `pages`. Each page binds `source`,
`geometry_hints`, and `recognition` by path/SHA-256. The recognized transcript records
its source binding and actual backend. Register this report as `offline-preparation`
and pass it to final QA with `--offline-preparation`; it must not be handwritten.

## Deck revision authority

`revision-checkpoints.json` binds the immutable original deck, the versioned revised deck, the requested edits, and the evidence used to accept both changed and intentionally unchanged content:

```json
{
  "schema_version": 1,
  "original": {"path": "/abs/original.pptx", "sha256": "..."},
  "revised": {"path": "/abs/revised-v2.pptx", "sha256": "..."},
  "checkpoints": [{
    "id": "rev-001", "page": 3, "object_id": "title-01",
    "change_type": "text",
    "claim_ids": ["claim-003"],
    "request": "Replace the title with approved wording",
    "expected_before": "Previous title",
    "expected_after": "Approved new title",
    "status": "passed",
    "evidence": [{"path": "/abs/qa/revision/page-03-before-after.png", "sha256": "..."}]
  }],
  "unchanged_checks": [{
    "page": 4, "fields": ["title", "42%", "Source: src-002"],
    "status": "passed", "evidence": [{"path": "/abs/qa/revision/page-04.json", "sha256": "..."}]
  }],
  "contact_sheet": {"path": "/abs/qa/revision/revised-contact-sheet.png", "sha256": "..."}
}
```

The original and revised paths must be distinct existing PPTX files with current hashes. Every requested page/object change needs exactly one checkpoint with a unique `(page, object_id)`, `change_type`, explicit `claim_ids` (or a nonempty `claim_not_applicable_reason`), a machine-checkable `expected_after`/`new_value`, and hash-bound `passed` evidence. Revision claim IDs must resolve to the claim ledger, reciprocally record that page in `used_on_slides`, and expose wording found in the revised native page via `slide_evidence` or the claim statement. `change_type: text` also requires `expected_before`/`old_value`; `object_id` binds an actual stable OOXML text shape (`shape-<cNvPr id>`, `id:<id>`, `name:<name>`, or the `title` alias), and before/after text is checked on that object rather than anywhere on the page. The `title` aliases imply `title_changed`; otherwise set `title_changed: true` explicitly. Set `notes_changed: true` on a checkpoint before changing native speaker notes. For `visual`, `geometry`, or `media`, record the expected result as a description and add `before_slide_fingerprint` and `after_slide_fingerprint` from the validator's semantic slide inventory; a `media` checkpoint must also change the actual embedded-picture fingerprint inventory. Several objects on one page are valid. Skipped, pending, duplicate page/object pairs, self-attestation, or unbound checkpoints fail. `unchanged_checks` must cover critical content outside the requested edit scope, and the validator compares semantic slide fingerprints to catch unrequested geometry, style, relationship, and media changes even when native text is unchanged. Masters/layouts/themes may change only with `allow_global_changes: true` plus hash-bound `global_change_evidence`. Final non-structural QA supplies the original deck, revised deck, and `revision-checkpoints.json` together; validating only the revised package is insufficient.

## Speaker notes and QA

`speaker-notes.json` is the notes authority and may be rendered to DOCX or embedded PPT notes:

```json
{
  "schema_version": 1,
  "total_seconds": 600,
  "timing": {
    "requested_seconds": null,
    "cjk_chars_per_minute": 260,
    "words_per_minute": 150,
    "sentence_pause_seconds": 0.35
  },
  "reference_authorities": {
    "slide_copy_ledger": {"path": "/abs/slide-copy-ledger.json", "sha256": "..."},
    "reference_deck": {"path": "/abs/final-deck.pptx", "sha256": "..."}
  },
  "slides": [{
    "page": 1, "title": "Exact title", "target_seconds": 30,
    "goal": "...", "script": "Natural spoken script",
    "cue": "Point to ...", "transition": "Next ...",
    "source_caution": "Use only the verified wording below", "claim_ids": ["claim-001"]
  }]
}
```

For final QA, bind notes to both the current `slide-copy-ledger.json` and the accepted reference deck. Pass the ledger and reference deck as top-level validator inputs even when `reference_authorities` is present. Page count, one-based order, and normalized titles must agree across all three authorities; notes may not add a factual claim absent from the claim ledger or silently follow an obsolete outline.

`timing` is optional for existing ledgers; absent settings use the estimator defaults
shown above. `requested_seconds` is null/absent when the user has not specified time.
Otherwise persist the same target with state `record --preference
requested_duration_seconds=<seconds>` so a self-consistent notes file cannot replace
the user's actual requirement. Each slide may add `pause_seconds` for intentional
demos or silence. The estimator counts the first nonempty `script`/`spoken_script`/
`speaker_notes` field and a separate spoken transition; metadata is excluded.
Run `speech_timing.py` before rendering and adjust wording or realistic timing.
Final QA also runs this estimator for both DOCX and embedded PPT notes. A report of
estimated time must be identified as an estimate, not an actual rehearsal duration.

When `ppt-notes` is requested, the notes-bearing PPTX is a validated deliverable, not merely a rendering convenience. Inspect the PPTX package's slide-to-`notesSlide` relationships and native notes text, then reconcile every embedded page against `speaker-notes.json`, `slide-copy-ledger.json`, and the reference deck. Missing/extra notes, notes attached to the wrong slide, empty embedded notes, or a JSON/DOCX-only pass while the PPTX lacks matching notes are hard failures.

## Render evidence

`--report-render-evidence`, `--image-pptx-render-evidence`, and
`--speaker-notes-render-evidence` accept a JSON manifest, not a loose image or
directory. The manifest binds the exact delivered artifact to every ordered,
fully decoded render page. Use this schema:

```json
{
  "schema_version": 1,
  "tool": "ppt-gen.render-evidence",
  "passed": true,
  "artifact": {"path": "/abs/report.docx", "sha256": "..."},
  "renderer": {
    "name": "documents:documents",
    "version": "current-runtime",
    "input_sha256": "..."
  },
  "artifact_page_count": 2,
  "pages": [
    {"page": 1, "path": "/abs/renders/page-01.png", "sha256": "..."},
    {"page": 2, "path": "/abs/renders/page-02.png", "sha256": "..."}
  ]
}
```

For a notes DOCX, `artifact_page_count` is the physical rendered-document page
count, not the number of slide sections. A cover, multiple short sections on
one physical page, or one long section flowing across pages is valid; section
count/order is checked separately against the notes ledger and reference deck.

`renderer.name` must be the companion/office renderer that actually opened the
artifact; record its version or implementation and repeat the exact artifact
hash as `input_sha256`. Page numbers must be contiguous and one-based. `artifact_page_count` must equal
the ordered array length and, when the document/deck exposes a deterministic
page count, that count as well. Every page image must have a short edge of at
least 720 px and a long edge of at least 1280 px, so both landscape slides and
portrait report pages are accepted, and must have a unique current SHA-256.
Final validation independently re-renders the bound DOCX/PPTX with the approved
local office/PDF renderer and compares each ordered page to the declared image;
renderer metadata alone is not trusted. A manifest bound to a different DOCX/PPTX, a
naked PNG/JPEG, a render directory, or a blank/near-uniform page is invalid final evidence.

`qa-report.json` records the result rather than only a Boolean:

```json
{
  "schema_version": 1,
  "tool": "ppt-gen.validate_delivery",
  "validation_contract_version": 2,
  "mode": "final-delivery",
  "status": "passed_with_warnings",
  "passed": true,
  "checked_at": "2026-08-12T12:00:00+08:00",
  "authorities": {
    "source-manifest": {"path": "/abs/source-manifest.json", "sha256": "..."},
    "claim-ledger": {"path": "/abs/claim-ledger.json", "sha256": "..."}
  },
  "state_checkpoint": {"path": "/abs/ppt-gen-state.json", "state_revision": 12, "verify_history_id": 12, "verify_revision": 12, "drift_count": 0},
  "artifacts": [{"role": "editable-pptx", "path": "/abs/deck.pptx", "sha256": "..."}],
  "gates": [{"id": "editable.page-01", "status": "passed", "evidence": ["validation.json", "render.png"], "messages": []}],
  "errors": [], "warnings": ["One non-blocking font substitution"]
}
```

Allowed status values are `passed`, `passed_with_warnings`, and `failed`. Any hard gate failure makes the overall result `failed`.

## Final-delivery validation bundle

A final, non-`--structural-only` validator run is a single evidence bundle, not a collection of independent artifact checks. It always includes top-level `source-manifest.json` and `claim-ledger.json`, the exact requested deliverable, and every authority/evidence file applicable to that deliverable. In particular:

- image and editable decks include `slide-copy-ledger.json` and `slide-manifest.json`; editable decks also include `editable-handoff.json`, the guarded offline-preparation execution report, and per-page/final companion validation;
- deck revisions include original PPTX, revised PPTX, and `revision-checkpoints.json`;
- speaker-notes DOCX includes `speaker-notes.json`, `slide-copy-ledger.json`, and the accepted reference deck;
- `ppt-notes` includes the notes-bearing PPTX plus the same notes, copy, and reference-deck authorities, and requires native embedded-note inspection.

If any applicable member is absent, malformed, stale, hash-mismatched, or inconsistent, final QA fails closed. `--structural-only` is reserved for parser/package diagnosis. It may report structural observations but cannot create final acceptance evidence, satisfy a phase gate, mark a deliverable complete, or authorize delivery; label its report `mode: structural-only` and retain a prominent warning.

## Validation rules

- Page arrays are contiguous, one-based, unique, and agree across all ledgers and decks.
- Titles, names, dates, values, units, and `critical_fields` must match native slide text or an explicitly approved raster transcription exactly after whitespace normalization.
- A raster `text_verified` attestation is valid only when bound to the current image hash and backed by readable-size visual/OCR review evidence; it is not inferred from repeating ledger fields.
- Every path exists and every stored hash is rechecked before resume and delivery.
- A contract with an unknown schema version, malformed JSON, missing required key, or duplicate page/ID fails closed.
- Non-structural final QA rejects missing top-level source/claim authorities, incomplete applicable bundles, or evidence validated in unrelated runs without current hash binding.
- A deck revision passes only when original, revised, and revision checkpoints agree and the original was not overwritten.
- Speaker notes pass only when the ledger, rendered notes format, slide-copy ledger, and reference-deck order/titles agree; requested `ppt-notes` additionally requires matching native embedded notes.
- Record worker implementation, OCR/privacy policy, source hashes, render evidence, and validation paths so a resumed run does not depend on chat memory.
