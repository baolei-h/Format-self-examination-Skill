---
name: bioactive-surface-book-auditor
description: Audit chapters and delivery packages for the Chinese monograph 《生物活性表面可控构建技术》 against the project's unified writing and formatting standard. Use when a user asks Codex/WorkBuddy/豆包/ChatGPT to 自查、校核、验收、格式检查、参考文献检查、图表检查、近5-10年进展检查、交稿前检查, or to safely normalize a chapter manuscript. Best input is a .docx chapter; .md/.txt can receive content-only checks. Default to audit-only; modify files only when explicitly requested.
---

# 《生物活性表面可控构建技术》章节自查

## Core rule

Treat `references/book-writing-standard.md` as the house standard unless the user supplies a publisher template or newer explicit requirement. Publisher instructions always override the house standard. Never silently replace a user-supplied requirement with a generic publishing convention.

Default to **audit-only**. Do not edit the manuscript unless the user explicitly asks to fix/normalize it.

## Workflow

1. Identify the input.
   - Prefer `.docx` chapter manuscripts.
   - For `.md`/`.txt`, perform content, hierarchy, terminology, citation, and reference checks; clearly mark page/font/image-DPI checks as unavailable.
   - If a delivery folder is available, also audit required subfolders and source files.

2. Load the standard.
   - Read `references/book-writing-standard.md` for exact formatting and editorial rules.
   - Read `references/audit-rubric.md` for severity, verdict, and report format.
   - Read `references/sample-prompts.md` only when invocation examples are useful.

3. Run deterministic checks for DOCX.
   - Execute:
     `python scripts/audit_docx.py <chapter.docx> --markdown <report.md>`
   - If a chapter delivery directory exists, add:
     `--package-dir <chapter-folder>`
   - If execution is unavailable, reproduce the same checks manually and state which checks could not be measured exactly.

4. Perform semantic/editorial checks that scripts cannot reliably judge.
   - Check whether each section is organized as a monograph narrative, not student-by-student thesis summaries.
   - Check the recommended logic: scientific problem -> principle -> construction method -> representative research -> mechanism -> application/limitations -> recent progress.
   - Match causal language to evidence strength. Flag unsupported wording such as “证实/决定/直接调控” when only correlation is shown.
   - Check terminology and abbreviation consistency.
   - Check whether figures, tables, data, and third-party material are traceable and whether copyright/permission risk is documented.
   - Check scientific image-integrity risks: missing scale bars, inconsistent processing, suspicious selective enhancement, or untraceable image reuse.
   - For Chapters 2-10, confirm a dedicated recent-progress/development-direction section exists and uses recent 5-10 year literature as the main update layer while retaining necessary classic papers.
   - Do not fabricate missing citations, DOI, data, permissions, or source files.

5. Reconcile script findings with semantic findings.
   - Do not mark a rule as failed when it is only a recommendation; label it `建议优化`.
   - Distinguish `可自动修复` from `需作者判断` and `需补充材料`.
   - If the script cannot resolve inherited Word styles, inspect the visible document or style definitions before declaring a font/size failure.

6. Return the audit report using the exact structure in `references/audit-rubric.md`.

7. If the user explicitly requests fixes:
   - Create a copy named `<original_stem>_规范修订版.docx`.
   - Only auto-fix safe formatting: A4/page margins, paragraph alignment/spacing/indent, heading/caption/reference styles, page number placement, simple numbering consistency, and other non-scientific layout issues.
   - Do **not** auto-change scientific claims, numeric data, statistical tests, reference metadata, figure pixels, biological labels, copyright statements, or causal interpretations.
   - Preserve the original file.
   - Re-run the audit after editing and report remaining issues.

## Required checks

Always cover these categories when the input permits:

- Page size, orientation, margins, header/footer, page number.
- Chapter/section hierarchy and numbering; no more than four heading levels; no more than 15 first-level sections per chapter.
- Chapter-length target: Chapter 1 approximately 5,000-10,000 Chinese characters; Chapters 2-10 approximately 50,000-60,000 as planning targets.
- Chinese/Latin fonts, font sizes, alignment, paragraph spacing, 1.5-line body spacing, 2-character first-line indent.
- Auto-generated TOC principles when a TOC is present.
- Figure/table/equation numbering by chapter; figure title below, table title above; text cites figure/table before appearance.
- Figure source/editability, raster effective DPI, scale bars, panel labels, grayscale/color accessibility, source/permission notes.
- Editable Word tables, three-line-table preference, units in headers, table notes.
- Editable equations and chemical structures; symbol definitions; ChemDraw/source-file expectations.
- Statistical reporting: n, repeat type, expression, test, error bars, significance legend.
- GB/T 7714-2025 sequential-number reference style; chapter-local numbering; citation-reference one-to-one mapping; DOI retention where available.
- SI/legal units, spacing between numbers and units, effective digits, symbols, italics/roman conventions.
- Abbreviation first-use expansion and terminology consistency.
- Thesis-source rewriting, source traceability, copyright/permission, research integrity, AI-assisted-content responsibility.
- Delivery package completeness and naming/version rules when a package folder is supplied.

## Platform portability

Keep the workflow filesystem-first and connector-free. This allows use in Codex, WorkBuddy, 豆包工作模式, and similar agents.

- If native Skills are supported, invoke this skill by name.
- If native Skills are not supported, load `SKILL.md` as the task instruction and keep the `references/` and `scripts/` paths adjacent.
- If Python or `python-docx` is unavailable, perform manual checks from the reference standard and explicitly label deterministic checks as `未自动测量`.

## Output discipline

- Respond in Chinese by default.
- Be exact and editorial, not encouraging or promotional.
- Cite exact paragraph/section/page identifiers from the manuscript whenever available.
- Do not output a numerical “quality score” unless the user asks for one.
- End with a short prioritized repair list: `交稿前必须修` / `建议统一修` / `可后置到出版社排版`.
