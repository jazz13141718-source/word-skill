---
name: word-skill
description: Install, uninstall, convert, format, reformat, audit, and visually verify Chinese or English DOCX deliverables from TXT, Markdown, DOCX, DOC, RTF, or ODT in any command-capable agent environment using built-in document presets, custom saved presets, optional reference-DOCX style matching, Pandoc/LibreOffice when needed, python-docx post-processing, separated build/final output directories, style normalization, inline citation superscripting, Chinese and English caption handling, table layout adjustment, strict quality gates, and render QA. Use when asked to export AI conversation content to Word, format reports, briefings, speeches, papers, meeting materials, proposals, existing Word files, audit a Word deliverable, package this workflow for another agent, or reformat one document according to a Word template.
---

# word-skill

Use this skill or command-line toolkit to turn Chinese or English AI-generated content or existing files into polished Word deliverables. It is not limited to papers and does not require a specific agent runtime.

For non-OpenAI/Codex agents, read `agents/generic.md` and call `scripts/format_docx_document.py` directly. `agents/openai.yaml` is optional OpenAI/Codex UI metadata.

Supported inputs:

- Text sources: `.md`, `.markdown`, `.mdown`, `.txt`
- Existing Word: `.docx`
- LibreOffice-convertible sources: `.doc`, `.rtf`, `.odt`

## Requirements

Runtime requirements:

- Python 3.10 or newer.
- `python-docx`, required for all formatting, reformatting, auditing, and table styling.
- Pandoc, required only for `.md`, `.markdown`, `.mdown`, and `.txt` sources.
- LibreOffice, required for `.doc`, `.rtf`, and `.odt` conversion, and required for render QA unless `--skip-render` is used.
- `pdftoppm`, required for PNG page renders unless `--skip-render` is used. It can come from Poppler or TeX Live; full TeX Live is not required if Poppler already provides `pdftoppm`, and Poppler is not required if TeX Live already provides `pdftoppm`.
- Fonts matching the selected preset. Chinese presets normally use `宋体`, `黑体`, `微软雅黑`, and `Times New Roman`; English presets normally use `Times New Roman` or `Arial`. If fonts are missing, the OS may substitute fonts; perform visual QA before delivery.
- Pillow is optional. If installed, the audit can estimate rendered-page ink density to catch blank or nearly blank pages more reliably.

Behavior requirements:

- Do not silently install system software or Python packages. Report missing requirements and wait for user approval before installing anything.
- Run `--check-env` before conversion or audit.
- Use `--strict` for user-facing deliverables unless the user explicitly asks for a draft.
- Treat `quality_audit.ok == true` as the minimum machine gate for delivery.
- Superscript inline numeric citations such as `[1]`, `[1-3]`, and `[1，2]` in the document body; do not superscript reference-list entries such as `[1] Author...`.
- Recognize Chinese captions such as `图 1-1` / `表 1-1` and English captions such as `Figure 1` / `Fig. 1` / `Table 1`.
- Inspect rendered pages manually when the output is important, especially pages with wide tables, images, formulas, references, or dense layouts.
- Deliver only the final DOCX from `docx_output/` or the explicit `--final` path. Do not deliver files from `_docx_style_build/`.
- Keep install/uninstall clean: install only the `word-skill` directory, and uninstall only a directory whose `SKILL.md` validates as `name: word-skill`.

## Install Or Remove

Install from a checked-out skill package:

```powershell
& "<python>" "<word-skill>/scripts/install_word_skill.py" --dry-run
& "<python>" "<word-skill>/scripts/install_word_skill.py"
```

Uninstall cleanly:

```powershell
& "<python>" "<word-skill>/scripts/uninstall_word_skill.py" --dry-run
& "<python>" "<word-skill>/scripts/uninstall_word_skill.py"
```

Use `--target "<skills-parent-or-word-skill-dir>"` for non-default locations. For generic agent setups, set `WORD_SKILL_HOME` or `AGENT_SKILLS_DIR` to a parent skills/tools directory or an exact `word-skill` target directory. Install and uninstall scripts only copy/remove the `word-skill` directory after validating its `SKILL.md`; they do not install system tools.

## Workflow

1. Identify the source file, target document type, and whether a reference Word file should control the style.
2. Run the environment check:

```powershell
& "<python>" "<word-skill>/scripts/format_docx_document.py" --check-env --source "<source-file>"
```

3. Format with a preset:

```powershell
& "<python>" "<word-skill>/scripts/format_docx_document.py" `
  --source "<source-file>" `
  --preset report `
  --strict
```

4. Or format according to a reference Word file:

```powershell
& "<python>" "<word-skill>/scripts/format_docx_document.py" `
  --source "<source-file>" `
  --style-source "<reference.docx>" `
  --preset reference `
  --strict
```

5. For an existing DOCX that should not be changed yet, audit only:

```powershell
& "<python>" "<word-skill>/scripts/format_docx_document.py" `
  --source "<existing.docx>" `
  --audit-only `
  --strict
```

6. Deliver only the final DOCX from `docx_output/` after `run_report.json` has `"quality_audit": {"ok": true, ...}`. Treat `_docx_style_build/<run-id>/` as internal diagnostics containing converted sources, processed text, Pandoc DOCX, render PDF/PNGs, and `run_report.json`.

## Presets

Built-in presets:

- `auto`: infer from filename/title; uses `reference` when `--style-source` is provided.
- `general`: ordinary AI-exported documents and mixed materials.
- `academic-paper`: papers and thesis-like manuscripts.
- `report`: formal reports and longer written materials.
- `briefing`: concise汇报材料/简报-style documents.
- `speech`: speeches, remarks, and speaking scripts.
- `english-general`: ordinary English AI-exported documents and mixed materials.
- `english-academic`: English papers, articles, manuscripts, and thesis-like documents.
- `english-report`: English reports and formal business/technical documents.
- `english-briefing`: concise English briefings and executive summaries.
- `english-speech`: English speeches, remarks, and scripts.
- `reference`: preserve styles/page layout from `--style-source`.

Custom presets:

```powershell
& "<python>" "<word-skill>/scripts/format_docx_document.py" `
  --source "<source-file>" `
  --preset report `
  --preset-file "<custom-format.json>" `
  --save-preset-name "my-report-style"
```

Saved presets are stored under `assets/custom_presets/` and can later be used with `--preset my-report-style`.

## Output Rules

- Final deliverables go only to `docx_output/` by default, or to the explicit `--output-dir` / `--final` path. Chinese presets default to `<source>_格式化.docx`; English presets default to `<source>_formatted.docx`.
- Intermediate files go only to `_docx_style_build/<timestamp_pid>/`.
- Do not hand users `_pandoc.docx`, `processed_paper.md`, converted source DOCX files, render PDFs, render PNGs, or build logs unless they explicitly ask for diagnostics.
- The script overwrites the deterministic final DOCX by default. Use `--no-overwrite` when preservation is needed.
- Keep final files and intermediate files separated when packaging for another agent. Clean test artifacts before sharing the skill package.

## Dependency Policy

After Python and `python-docx` are available, the formatter's `--check-env` mode checks for:

- Pandoc: required for `.md`, `.markdown`, `.mdown`, and `.txt` input.
- LibreOffice: required for `.doc`, `.rtf`, `.odt` input and for render QA.
- pdftoppm: required for PNG page renders. It can come from Poppler or from a TeX Live installation; users do not need full TeX Live if Poppler already provides `pdftoppm`, and they do not need Poppler if TeX Live already provides `pdftoppm`.

Do not silently install system software. If a dependency is missing, report the missing tool and give the install command or download link from the script output. Only install tools after the user explicitly asks.

Read `references/environment-and-output.md` when troubleshooting dependencies, output directories, presets, reference-style matching, or table styling.

## Quality Review Loop

Use `--strict` for user-facing deliverables. The script writes a machine-readable `quality_audit` with:

- `ok`: true only when no blocking findings remain.
- `findings`: blocking issues such as corrupt DOCX packages, missing images, BodyText style drift, non-superscript inline citations, blank rendered pages, missing renders, malformed tables, or final output placed in a build directory.
- `warnings`: non-blocking issues requiring human attention, such as wide tables, external images, skipped render QA, or density checks unavailable without Pillow.
- `metrics`: document size, paragraph/table counts, render page count, and sampled render density.

The formatter runs a small repair loop controlled by `--qa-iterations` (default `2`) for repairable issues such as BodyText style residue and table geometry. If `quality_audit.ok` is false after the loop, inspect `findings`, adjust the source/preset/reference, and rerun before delivering.
