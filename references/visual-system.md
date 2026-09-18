# Visual direction and image generation

## Generate comparable style previews

Use built-in `imagegen`, one call per candidate, on identical representative content. Vary the whole system—not only colors:

- palette and contrast;
- Chinese typography and hierarchy;
- grid, density, whitespace, and rhythm;
- photo/illustration treatment;
- icon and shape language;
- chart styling and source footers;
- ease of object-level reconstruction.

Use real project titles and a small set of ledger-approved copy. Save prompts and output hashes. Populate `style-options.json` immediately after each candidate so a stop does not erase the decision context.

## Score candidates in continuous mode

Reject wrong ratio, false/garbled critical copy, illegible contrast, or a language that cannot extend across all six layout types. Score survivors:

- Chinese/text readability and accuracy: 30%
- audience/topic fit: 25%
- data/comparison communication: 20%
- cross-slide consistency/extensibility: 15%
- editable reconstruction simplicity: 10%

Choose the highest score; break ties with clearer hierarchy and simpler geometry. Persist `selected_id`, scores, and reason. In guided/stepwise mode, do not advance from `style-options` until `selected_id` is populated.

## Lock a deterministic visual system

Translate the selected candidate into `visual-system.json` before making the template overview. Bind it to the absolute `style-options.json` path/hash and the same `selected_id` recorded in state. At minimum lock palette, font families/weights, 1920×1080 geometry, margins/grid/gutters, imagery, icons, charts, footer/source convention, and these layout IDs:

1. `cover`
2. `section`
3. `narrative`
4. `comparison`
5. `data-chart`
6. `conclusion`

Generate each layout as its own exact-16:9 page. Validate all six separately, then compose the overview deterministically:

```bash
python <skill-root>/scripts/compose_template_overview.py \
  --visual-system /absolute/visual-system.json \
  --output /absolute/template-overview.png \
  --slide "封面=/absolute/01-cover.png" \
  --slide "章节=/absolute/02-section.png" \
  --slide "叙事=/absolute/03-narrative.png" \
  --slide "对比=/absolute/04-comparison.png" \
  --slide "数据=/absolute/05-data-chart.png" \
  --slide "结论=/absolute/06-conclusion.png"
```

The composer—not the generative model—owns the 3×2 geometry. Keep the overview 1920×1080 and every source thumbnail exact 16:9.

## Layout-page prompt contract

Use this for each of the six source layouts, substituting only ledger-authorized copy:

```text
Use case: productivity-visual
Asset type: one Chinese PPT layout specimen, exact 16:9
Layout role: <cover|section|narrative|comparison|data-chart|conclusion>
Style system: <visual-system.json summary>
Exact text: <verbatim title/body/labels/values/units/footer>
Content: focused but sufficiently rich to demonstrate hierarchy; useful whitespace; no filler
Continuity: lock margins, title zone, footer, colors, icons, imagery, and chart grammar
Constraints: no frame/contact sheet; no watermark; no invented text/data/logo/quote/source; all critical copy readable
```

## Per-slide prompt contract

For every final raster slide supply:

```text
Use case: productivity-visual
Asset type: final Chinese presentation slide, exact 16:9
Page identity: <page/id/title from slide-copy-ledger.json>
Primary request: <one central judgment and page role>
Style system: <selected visual-system.json>
Exact text: <title, copy, labels, values, units, source footer verbatim>
Visual structure: <comparison/chart/timeline/matrix/photo-led composition>
Content density: rich enough to support the argument, with clear hierarchy and useful whitespace
Continuity: match approved layout in margins, title zone, footer, color, icons, and chart grammar
Constraints: no watermark; no invented text/data/logo/quote/source; all critical copy readable
```

Issue one built-in call per distinct page. Copy project-bound results into the project; do not rely on the global generated-image path.

## Correct and register each page

Treat generated output as visual production, never as factual authority:

1. Fully decode the file and normalize to exact 16:9 with non-distorting crop/pad.
2. Reconcile titles, names, dates, numbers, units, charts, and footers with `slide-copy-ledger.json`.
3. Regenerate/edit any false, garbled, clipped, or unreadable critical content.
4. Add the accepted path/hash/dimensions/critical fields to `slide-manifest.json`.
5. Inspect a natural-order contact sheet for rhythm and every page at readable size.

AI typography may later be rebuilt as native editable objects, but the requested image deck must itself remain correct and readable.

## Reuse approved templates and unchanged pages

When the user supplies or approves an existing complete template package, reuse its
selected style ledger, visual system, and six layout files after checking hashes,
geometry, readability, font availability, and suitability for this project. Preserve
its original provenance; do not claim old previews were generated for new content.
Do not invent two or three new choices when the user has already selected a style.
If a new overview is requested, compose all six validated layouts as usual.

After accepting a raster deck, record a versioned per-page reuse index. The QA report
must bind the current copy ledger, slide manifest, and visual system:

```bash
python <skill-root>/scripts/page_reuse.py record \
  --ledger /abs/slide-copy-ledger.json --visual-system /abs/visual-system.json \
  --slide-manifest /abs/slide-manifest.json --qa-report /abs/qa-report.json \
  --index /abs/reuse-index-v1.json
python <skill-root>/scripts/page_reuse.py plan \
  --ledger /abs/slide-copy-ledger.json --visual-system /abs/visual-system.json \
  --index /abs/reuse-index-v1.json
```

The planner compares complete page-copy snapshots, stable IDs/page numbers, selected
visual-system hash, original QA hash, and image bytes. Only unchanged pages become
`reuse-candidate`; they are not automatically accepted. Regenerate changed pages,
reconcile reused pages to current sources/claims, refresh state checkpoints, and run
whole-deck final QA. Global design changes invalidate all pages. Never reuse a full
source-slide raster as an editable reconstruction artifact.
