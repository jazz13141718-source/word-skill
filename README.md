# word-skill

`word-skill` is a Codex skill for converting, formatting, auditing, and visually checking Chinese or English Word deliverables. It supports Markdown, plain text, DOCX, DOC, RTF, and ODT sources, with separated final/build outputs and machine-readable quality reports.

This repository is public source / source-available under a non-commercial license. Commercial use is not allowed without separate written permission. Derivative works are not required to use the same license or publish their complete source code, as long as their use and distribution remain non-commercial and preserve required attribution.

## What It Does

- Convert `.md`, `.markdown`, `.mdown`, and `.txt` to DOCX through Pandoc.
- Reformat existing `.docx` files with consistent page, style, table, footer, and citation handling.
- Convert `.doc`, `.rtf`, and `.odt` through LibreOffice.
- Match another Word file's style with `--style-source`.
- Use Chinese or English presets, including `english-general`, `english-academic`, `english-report`, `english-briefing`, and `english-speech`.
- Recognize Chinese captions such as `图 1-1` / `表 1-1` and English captions such as `Figure 1` / `Fig. 1` / `Table 1`.
- Keep final DOCX files separate from `_docx_style_build` diagnostics.
- Audit output through `quality_audit`, including BodyText style drift, missing images, malformed tables, render output, blank pages, and inline citation superscripts.

## Requirements

- Python 3.10 or newer.
- Python packages in `requirements.txt`.
- Pandoc for Markdown/text input.
- LibreOffice for DOC/RTF/ODT conversion and render QA.
- `pdftoppm` for PNG page renders. It can come from Poppler or TeX Live.
- Fonts required by the selected preset. Chinese presets commonly use SimSun, SimHei, Microsoft YaHei, and Times New Roman. English presets commonly use Times New Roman or Arial.

The skill does not silently install system tools. Run `--check-env` first and install missing tools only after user approval.

## Install

From this repository root:

```powershell
& "<python>" ".\word-skill\scripts\install_word_skill.py" --dry-run
& "<python>" ".\word-skill\scripts\install_word_skill.py"
```

Install to a custom skills directory:

```powershell
& "<python>" ".\word-skill\scripts\install_word_skill.py" --target "D:\SomeAgent\skills"
```

Uninstall:

```powershell
& "<python>" ".\word-skill\scripts\uninstall_word_skill.py" --dry-run
& "<python>" ".\word-skill\scripts\uninstall_word_skill.py"
```

## Usage

Check the environment:

```powershell
& "<python>" ".\word-skill\scripts\format_docx_document.py" --check-env --source ".\paper.md"
```

Format a report:

```powershell
& "<python>" ".\word-skill\scripts\format_docx_document.py" `
  --source ".\paper.md" `
  --preset report `
  --strict
```

Format an English report explicitly:

```powershell
& "<python>" ".\word-skill\scripts\format_docx_document.py" `
  --source ".\report.md" `
  --preset english-report `
  --strict
```

Audit an existing DOCX without changing it:

```powershell
& "<python>" ".\word-skill\scripts\format_docx_document.py" `
  --source ".\document.docx" `
  --audit-only `
  --strict
```

Use another Word file as the style source:

```powershell
& "<python>" ".\word-skill\scripts\format_docx_document.py" `
  --source ".\target.docx" `
  --style-source ".\reference.docx" `
  --preset reference `
  --strict
```

## Output Policy

- Final DOCX files go to `docx_output/` by default, or to `--output-dir` / `--final`.
- Chinese presets default to `<source>_格式化.docx`; English presets default to `<source>_formatted.docx`.
- Build artifacts go to `_docx_style_build/<run-id>/`.
- Deliver only the final DOCX unless diagnostics are explicitly requested.
- For important deliverables, inspect rendered PNG pages before delivery.

## Development

Install Python dependencies:

```powershell
& "<python>" -m pip install -r requirements.txt
```

Run the smoke tests:

```powershell
& "<python>" tests\test_word_skill_smoke.py
```

The smoke test validates required skill files, compiles bundled scripts, and checks that inline body citations are superscripted while `[0,1]` intervals and reference-list entries are left alone.

## License

`word-skill Non-Commercial License 1.0`. See `LICENSE`.

This is not an OSI-approved open source license because it restricts commercial use. It is intended for public, non-commercial collaboration.
