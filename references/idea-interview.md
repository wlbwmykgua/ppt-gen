# Idea interview: grill-me before production

Read this for a new brief, an unresolved inherited brief, or a material change of direction. The purpose is a better agreed plan, not a longer questionnaire. Apply it only to the user's requested outputs and unresolved decisions.

## Resolve the companion and the user's mode

- Load the installed `grill-me` skill completely, then its `grilling` implementation completely. The shim's “Skill tool” means the host's available skill-loading mechanism; if no literal Skill tool exists, read the resolved `SKILL.md` files. Do not invent a tool call or copy a personal absolute path into the workflow.
- If only `grilling` is installed, use it directly. If neither is available, disclose that briefly and use the rounds/confirmation procedure below as Ppt Gen's local fallback. Do not silently install anything or claim the companion ran.
- Default: interview first. Explicit “直接做 / 不要问 / 其他你定 / 全套自动完成” delegates the unresolved choices in the scope of that instruction; record them as assumptions, not user-selected answers. An answer delegating one question does not delegate the entire project.
- “先问清楚，再自动做完” means interview now, then `continuous` production after confirmation. A later instruction to stop asking can end the interview and delegate remaining nonessential choices. Privacy/authorization blockers still require resolution.
- On an unchanged resumed project, reuse the accepted brief; do not run a fresh interview just because the user says “继续”. Do not reinterpret an old confirmed brief as unapproved solely because this feature was added.

## Work the decision tree in rounds

Extract existing answers and inspect supplied material before asking. Distinguish observed facts, user choices, unresolved choices, and explicitly delegated defaults. A recommendation is not an answer until accepted or delegated.

Build a small tree of decisions relevant to this task. Its frontier is all unresolved decisions whose prerequisites are settled. Ask that whole frontier in one round; questions that depend on an unanswered question belong to a later round. Recompute after each answer. Do not impose a three-question/two-follow-up cap or ask every possible question upfront.

Use the companion's numbered question/recommendation format in the user's language, with a short reason or tradeoff:

```text
❓ **Q1** - **受众与效果**：这份介绍主要给谁看，希望他们看完改变哪一个认识？

➡️ 建议：如果是面向不了解该人物的同学，以“理解其独特之处”为目标，比罗列履历更容易记住；若你的用途不同，我们再调整。
```

Present the round, then wait for the user's answers. Do not start deliverable generation while waiting. If an answer is incomplete, keep only the material unresolved branch open; “不知道，你建议” is an opportunity to explain options, not to silently choose unless the user delegates.

Useful branches (not a mandatory form):

- **Outcome → audience → angle:** What should this audience understand, decide, or do? What do they already know? Which single thesis or question organizes the material?
- **Scope → content → evidence:** What is in/out; which examples, comparisons, counterarguments, or misconceptions matter? What counts as convincing evidence, and what source/cutoff constraints apply? Frame proposed claims as hypotheses until checked.
- **Deliverables → depth → structure:** Which outputs, start/end points, and acceptance criteria? Only then resolve appropriate length, information density, slide sequence, or speaking time. Surface conflicts such as “全面深入” versus “三分钟”, and offer a concrete tradeoff.
- **Audience/usage → tone and visual direction:** Formal, teaching, analytical, narrative, etc.; brand/template and dislikes. Set direction now, but keep the actual 2–3 generated style previews and selection at `style-options` after the outline. Do not generate sample images during the interview.
- **Sources/privacy → feasible route:** Preserve supplied work, decide whether updating it is allowed, and settle confidentiality before remote operations. Offline OCR remains the existing default, not a recurring token question.

Challenge weak assumptions respectfully: identify the tension, explain how it affects this audience, offer alternatives, then let the user decide. Avoid manufactured objections, unsupported factual claims, and aesthetic trivia that will be resolved by the later previews.

Look up discoverable facts rather than asking the user to do research. Follow `grilling`'s fact-exploration delegation when available and authorized; any standing requirement for approval before a new subagent still applies. State the subtask and write scope before seeking that approval. If delegation is unavailable or not authorized, disclose the limitation and use permitted read-only checks yourself. While a fact is unresolved, defer only its dependent questions; ask the rest of the frontier. Never upload confidential material to answer a question.

## Respect arbitrary entry and exit points

- Report-only: discuss purpose, depth, structure, evidence, and writing tone; no slide count, speaking duration, palette, or extra deck.
- Existing report/outline: inspect and preserve accepted content; refine only unresolved downstream presentation choices. A better story is not permission to rewrite source claims.
- Faithful image-to-editable conversion: settle page scope, fidelity/editability tradeoffs, and privacy; skip topic-development and redesign questions.
- Script-only: use the actual deck order and content. Ask only material delivery/audience questions; if no duration was requested, propose a natural estimate without making a time choice mandatory.
- Bounded revisions or QA-only: do not expand a precise task into a new presentation strategy. If the request already specifies and authorizes the complete plan, record that as confirmation and proceed; do not ask for duplicate approval.

## Close the interview and persist it

Stop when all material branches for this scope are settled or explicitly delegated. Summarize a compact **制作方案**: goal/audience, thesis and structure, evidence/source policy, requested files and range, constraints/design direction where relevant, assumptions, and production mode. Ask for confirmation before substantive research/writing, generated previews, or deck production. Do not keep inventing branches after agreement.

An explicit confirmation of this summary (including “按这个继续”) opens the production gate. A bare “继续” in the middle of unanswered rounds does not silently select every recommendation; continue the interview or clarify whether the user is delegating. A fully delegated automatic run needs no extra summary-confirmation turn: state assumptions briefly and proceed. This does not replace the later style-choice gate in guided/stepwise runs.

For resumable work, save `sources/production-brief.md` with the current summary, decision dependencies, answers and their provenance, open frontier, confirmation/delegation evidence, and latest round. Preserve `sources/user-brief.md` verbatim. Before state exists, the saved brief is enough; do not initialize invented production targets merely to store questions.

Once the route is settled, initialize or reuse normal project state. Use existing `record --preference` fields such as `intake_mode=interview|delegated`, `brief_status=interviewing|awaiting_confirmation|confirmed|delegated`, and `production_brief_path=/absolute/path`; these are descriptive metadata, not script-enforced gates. Record material answers with `--decision` and delegated defaults with `--assumption`. Use the dedicated flags for source/privacy policies, not arbitrary preference keys.

When awaiting answers, persist the brief and `pause` an initialized run with a reason; do not mark a production phase complete or add a mandatory `brief` target. At confirmation/delegation, register the production brief as an authority input owned by the earliest requested phase it affects, then `verify` and `resume`. On a later scope change, revisit only affected decisions, version the brief, and register its changed authority so dependent work is invalidated. Unaffected accepted work remains reusable.
