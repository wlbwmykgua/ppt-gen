# Intake, routing, and preflight

## Refine the idea before generation

Read [idea-interview.md](idea-interview.md) for the default `grill-me` → `grilling` interview and its confirmation gate. Infer answered items from the prompt and inspect supplied files first. Build questions around material unresolved decisions, not a fixed questionnaire:

1. **范围与运行方式**：从哪一步开始、做到哪一步；具体要哪些交付物；访谈结束后关键节点确认、每步确认，还是全自动不中断？
2. **场景与目标**：给谁看、用在哪里、希望观众理解或采取什么行动；据此再讨论论点、深度、结构和必要的时长/页数。
3. **内容、视觉与隐私**：必须出现/避免什么；证据与反方观点；主文件、品牌、模板、来源/截止日；允许标准云端图片生成，还是素材必须留在本机？

These are decision areas, not three mandatory questions. Ask the currently independent questions in rounds, each with a recommendation; wait for answers before dependent follow-ups. Never re-ask supplied information. End with a concise agreed plan and wait for confirmation before production. Honor explicit delegation such as “全套自动完成，其他你定”; record assumptions and proceed without nonessential questions.

Before confirmation, local input inspection and saving interview state are allowed. Targeted fact checks needed to frame a question follow the settled privacy policy; substantive report research/writing, style previews, rendering smoke artifacts, and deck generation wait for the production gate.

## Adapt every input before routing

For any supplied DOCX, Markdown, PDF, PPTX, image folder, or mixed set, run:

```bash
python <skill-root>/scripts/source_adapter.py <project-dir> \
  --input primary=/absolute/input/path \
  --source-policy faithful --data-handling standard
```

Inspect actual contents rather than extensions. Adopt the resulting `source-manifest.json`; do not regenerate valid upstream work. Preserve originals and hashes. If adapters initialize ledgers, review them before treating them as approved authority.

If the request is only a topic or one sentence, save that exact wording to `sources/user-brief.md` and use it as the initial local source. For researched web/PDF authorities, create project-local evidence records containing title, canonical URL, retrieval time, and the relevant locator/summary; hash and append those records to `source-manifest.json`. A chat message, search result ID, or uncaptured URL is not a resumable source authority.

Choose one source policy:

- `faithful`: preserve the supplied story/copy; verify only where needed to identify a material error or risk, and disclose conflicts instead of silently rewriting.
- `verify-update`: verify unstable claims with current authoritative sources and update with citations and a cutoff.

## Infer the start phase

| Input or request | Default start |
|---|---|
| Brief topic or loose notes | `report` |
| Substantive DOCX/PDF/Markdown report | `outline` |
| Complete page-by-page outline | `style-options` |
| Approved style/brand/template | `template` or `image-deck` |
| Slide images, scanned PDF, image-only PPTX | `editable-deck` |
| Existing editable PPTX plus revision | `deck-revision` |
| Existing PPTX plus script only | `speaker-notes` |
| Existing requested deliverable plus QA only | `qa-package` |

“从报告开始做 PPT” means the report is input, not permission to rewrite it. “只做报告” ends at `report`. A revision request is a real `deck-revision` phase, not an informal exception.

## Separate route from requested files

Record `targets` (execution phases) and `deliverables` (files) independently. If the user requests PNG slides but no image PPTX, target `image-deck` and request only `slide-images`. If the user supplies all content and asks only for a revised PPTX, target `deck-revision` only.

For existing deliverables plus QA only, import them with explicit validator roles;
do not request them as newly generated outputs or activate their production stages:

```bash
python <skill-root>/scripts/project_state.py init <project-dir> \
  --start-at qa-package --stop-after qa-package --source-policy faithful \
  --deliverable qa-report
python <skill-root>/scripts/project_state.py add-input <project-dir> \
  existing=/absolute/existing-report.docx --stage qa-package --artifact-role report
```

Register source/claim authorities and the usual render/evidence bundle, run `verify`,
then validate with `--report` and the full report gate. The imported artifact remains
untouched. For other kinds use the matching validator role (`image-pptx`,
`editable-pptx`, `speaker-notes`, etc.). An imported artifact needs the same substantive
QA as a generated one; missing evidence is not a reason to silently waive a gate.

## Run preflight before expensive work

Decide `data_handling` before any remote call, then run:

```bash
python <skill-root>/scripts/preflight.py \
  --project-dir <project-dir> --data-handling standard \
  --require-docx --require-pptx --require-editable
```

Use `local-only` when the user requires confidential/fully local processing. In that mode, do not call remote research or image generation and disable remote page-asset backends. Acknowledge any visual-fidelity tradeoff instead of violating privacy.

The preflight JSON must record app/CLI/font availability, OCR and image-backend policy, agent-only checks still required, warnings, and overall tri-state. Before a long production run, make and render a tiny Chinese DOCX/PPTX smoke artifact with the selected companion tools. Verify that `imagegen` and worker dispatch are actually callable when required; a successful CLI check alone is insufficient.

## Defaults

- Language: follow the user and primary material.
- Canvas: exact 16:9.
- Length: content-driven, commonly 10–12 slides and 12–15 minutes if unspecified; do not pad.
- Report: substantive and evidence-backed.
- Story grammar: comparison + data + charts only when reliable data supports them; otherwise timeline, matrix, flow, or qualitative framework.
- Source policy: `faithful` for supplied substantive material; `verify-update` for a bare topic requiring research.
- Data handling: ask or infer `standard`; never infer `local-only` away from an explicit confidentiality request.
- Style previews: three by default, two when speed matters.
- Editable conversion: `image-to-editable-ppt`, offline OCR/hints, no PaddleOCR token.
- Speaker notes: 45–90 seconds per content page and 20–40 seconds for a cover when time is unspecified.
- Files: isolated project directory, absolute paths, versioned outputs, no original overwrite.

## Skip irrelevant questions and phases

- Report-only: no questions about slide count, duration, or palette.
- Existing report → PPT: read it; interview only unresolved presentation decisions (audience, emphasis, narrative, length, hard constraints, privacy). Do not reopen accepted report claims or rewrite its argument without permission.
- Existing outline → PPT: no new report/research unless requested.
- Faithful image → editable PPT: skip style previews.
- Script-only: no palette question; infer duration from deck density.
- Supplied page count supersedes duration; supplied duration allows page count inference.
- Supplied template/brand is the highest visual constraint.

## Run-mode behavior

Interview depth and production run mode are separate. An explicit request for an interview followed by automatic production uses the interview gate first, then `continuous`. A plain request for uninterrupted automatic production delegates intake decisions as described in `idea-interview.md`.

- `guided`: pause at style choice and any material ambiguity.
- `stepwise`: persist and pause after every requested phase.
- `continuous`: record assumptions and proceed; never skip gates.

On a user stop, finish only the current atomic file write, persist page/checkpoint state, mark the run paused, and report the last accepted artifact plus next action. On “继续”, run `verify` first; changed/missing authorities become stale before resuming. If the user extends the endpoint, use `retarget` instead of reinitializing the project.
