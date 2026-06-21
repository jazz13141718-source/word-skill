# Environment, Output, And Style Policy

## Required Tools

Run the environment check first:

```powershell
& "<python>" "<word-skill>\scripts\format_docx_document.py" --check-env --source "<source-file>"
```

The script searches explicit arguments, environment variables, PATH, and common Windows install locations.

## Requirement Matrix

Core runtime:

- Python 3.10 or newer.
- `python-docx`, required by the formatter script before any DOCX work can run.
- Access to the `word-skill` directory containing `SKILL.md`, `scripts/`, `references/`, and `assets/reference.docx`.

Input-dependent tools:

- `pandoc`: required for `.md`, `.markdown`, `.mdown`, and `.txt` input.
- `libreoffice`: required for `.doc`, `.rtf`, and `.odt` input conversion. Also required unless `--skip-render` is used.
- `pdftoppm`: required unless `--skip-render` is used. Converts rendered PDF to page PNGs. It may be supplied by Poppler or TeX Live.

`.docx` input can be reformatted without Pandoc. If `--skip-render` is also used, `.docx` input does not require LibreOffice or pdftoppm.

Quality-enhancing requirements:

- Pillow is optional. When installed, rendered PNG pages are sampled for non-white pixel density so blank-page checks are stronger.
- Fonts used by the selected presets should be installed. Chinese presets commonly use `宋体`, `黑体`, `微软雅黑`, and `Times New Roman`; English presets commonly use `Times New Roman` or `Arial`. If a host system substitutes fonts, rendered QA and manual inspection become mandatory before delivery.
- Microsoft Word is not required for this pipeline. Word may still be useful for final human inspection when the user has it.

Render-tool choices:

- Use Poppler when the user wants the smallest dependency that provides `pdftoppm`.
- Use TeX Live when it is already installed or when the user also needs TeX/LaTeX tooling. Do not require a full TeX Live install solely for this skill if Poppler is acceptable.
- If both Poppler and TeX Live are installed, any working `pdftoppm` is acceptable. Prefer an explicit `--pdftoppm` path when multiple versions cause confusion.

Supported explicit overrides:

```powershell
--pandoc "C:\Program Files\Pandoc\pandoc.exe"
--libreoffice "D:\LibreOffice\program\soffice.com"
--pdftoppm "D:\texlive\2026\bin\windows\pdftoppm.exe"
```

Supported environment variables:

- `PANDOC`
- `LIBREOFFICE`
- `SOFFICE`
- `PDFTOPPM`

Python package installation examples, only after user approval:

```powershell
& "<python>" -m pip install python-docx
& "<python>" -m pip install Pillow
```

## Missing Tools

Do not install tools silently.

When tools are missing, report the missing dependency and provide these options:

- Python 3.10+: install from `https://www.python.org/downloads/`, use the agent's bundled Python, or use an existing compatible Python.
- `python-docx`: `"<python>" -m pip install python-docx`
- Pillow, optional: `"<python>" -m pip install Pillow`
- Pandoc: `winget install JohnMacFarlane.Pandoc`
- LibreOffice: `winget install TheDocumentFoundation.LibreOffice`
- pdftoppm render tool: install Poppler with `winget install oschwartz10612.Poppler`, or install/use an existing TeX Live distribution that includes `pdftoppm`.
- Chinese fonts: install the required fonts or switch to a preset/custom preset using fonts available on the host machine.

If the user has already installed a tool in a nonstandard location, rerun with the explicit path argument.

Never mark missing Pillow as a blocker by itself. Missing Pillow weakens blank-page detection but does not prevent formatting or rendering.

## Install And Uninstall

Install the skill for an agent that uses the default discovery location:

```powershell
& "<python>" "<word-skill>\scripts\install_word_skill.py" --dry-run
& "<python>" "<word-skill>\scripts\install_word_skill.py"
```

Install to a custom skills directory:

```powershell
& "<python>" "<word-skill>\scripts\install_word_skill.py" --target "D:\SomeAgent\skills"
```

Install with generic agent environment variables:

```powershell
$env:AGENT_SKILLS_DIR = "D:\SomeAgent\skills"
& "<python>" "<word-skill>\scripts\install_word_skill.py"

$env:WORD_SKILL_HOME = "D:\SomeAgent\skills"
& "<python>" "<word-skill>\scripts\install_word_skill.py"
```

Uninstall cleanly:

```powershell
& "<python>" "<word-skill>\scripts\uninstall_word_skill.py" --dry-run
& "<python>" "<word-skill>\scripts\uninstall_word_skill.py"
```

The scripts validate `SKILL.md` before replacing or deleting anything. They never install Pandoc, LibreOffice, Poppler, or Python packages automatically; they only report missing requirements so the user or host agent can decide how to install them. If no explicit target or generic environment variable is provided, the fallback is Codex-compatible for backward compatibility.

## Operating Requirements

- Preserve the user's source file. The formatter stages a copy in `_docx_style_build`; do not edit or overwrite the source unless the user explicitly asks.
- Use `--audit-only` when the user only wants a review of an existing DOCX.
- Use `--style-source "<reference.docx>" --preset reference` when one Word file should control another file's formatting.
- Use `--preset-file` for one-off custom formatting and `--save-preset-name` only when the user wants to keep a reusable preset.
- Use `--no-overwrite` when replacing an existing final DOCX would be risky.
- Use explicit tool paths when dependencies are installed outside PATH.
- Use English presets (`english-general`, `english-academic`, `english-report`, `english-briefing`, `english-speech`) for English-only documents when `auto` does not infer the desired style.
- Keep inline numeric citations in the body as superscripted bracketed numbers, for example `[1]`, `[1-3]`, or `[1，2]`. Do not superscript numbered entries in the final reference list.
- Keep captions consistent with the document language. Chinese captions such as `图 1-1` / `表 1-1` and English captions such as `Figure 1` / `Fig. 1` / `Table 1` are recognized and centered.
- Treat warnings as follow-up work for important deliverables. Wide tables, external images, skipped render QA, and missing Pillow require either visual inspection or user disclosure.
- Do not package test outputs, `_docx_style_build`, `docx_output`, or `__pycache__` when sharing the skill with another agent.

## Output Separation

Default output layout for a source file `<project>\document.docx`:

```text
<project>\
  docx_output\
    document_格式化.docx
  _docx_style_build\
    <YYYYMMDD_HHMMSS_pid>\
      processed_paper.md       # text/Markdown inputs only
      pandoc.docx              # text/Markdown inputs only
      source.docx              # DOCX inputs only
      lo_docx\                 # DOC/RTF/ODT inputs only
      formatted.docx
      run_report.json
      render_pdf\
      render_png\
```

Only files in `docx_output\` are final deliverables. Chinese presets default to `document_格式化.docx`; English presets default to `document_formatted.docx`.

## Presets

Use `--preset` for built-in or saved presets:

- `auto`
- `general`
- `academic-paper`
- `report`
- `briefing`
- `speech`
- `english-general`
- `english-academic`
- `english-report`
- `english-briefing`
- `english-speech`
- `reference`

Custom preset JSON can override any subset of the preset:

```json
{
  "page": {
    "top_cm": 2.5,
    "bottom_cm": 2.5,
    "left_cm": 2.8,
    "right_cm": 2.6
  },
  "fonts": {
    "east_asia": "宋体",
    "ascii": "Times New Roman"
  },
  "body": {
    "size": 11,
    "line": 1.45,
    "first_indent_cm": 0.78,
    "space_after": 2
  },
  "title": {
    "size": 20,
    "bold": true,
    "line": 1.2,
    "space_after": 16
  },
  "headings": {
    "h1": 15,
    "h2": 12,
    "h3": 11
  }
}
```

Use `--preset-file custom.json --save-preset-name name` to store a reusable preset under `assets/custom_presets/name.json`.

## Reference DOCX Matching

Use this when the user says to format one document according to another Word file:

```powershell
& "<python>" "<word-skill>\scripts\format_docx_document.py" `
  --source "<target-file>" `
  --style-source "<reference.docx>" `
  --preset reference
```

For text inputs, the reference DOCX is used as Pandoc's reference document. For DOCX-like inputs, key Word style parts such as styles, theme, font table, settings, and first-section page layout are imported from the reference file.

Reference matching preserves the reference style system as much as possible while still normalizing output folders, page numbers, body-style drift, and tables.

## Quality Gates

Before delivery:

1. Run formatting with `--strict` unless the user explicitly wants a best-effort draft.
2. Confirm `run_report.json` has `"quality_audit": {"ok": true}`.
3. Confirm `quality_audit.findings` is empty.
4. Confirm render page count is plausible.
5. Inspect representative PNG pages and pages near wide tables, figures, or references.
6. If the document is final for submission, inspect every page visually.

For an existing DOCX that should not be modified yet:

```powershell
& "<python>" "<word-skill>\scripts\format_docx_document.py" `
  --source "<existing.docx>" `
  --audit-only `
  --strict
```

The audit reports:

- blocking `findings`: corrupt DOCX, missing image parts, BodyText style drift, non-superscript inline numeric citations, blank rendered pages, missing render output, malformed tables, final files placed in `_docx_style_build`.
- non-blocking `warnings`: wide tables, external images, skipped render QA, or unavailable PNG density checks.
- `metrics`: file size, paragraph/table counts, section count, render page count, and sampled non-white pixel ratios.

The formatter uses `--qa-iterations` (default `2`) to retry repairable issues such as BodyText style residue and table geometry. If `quality_audit.ok` remains false, do not deliver the DOCX; fix the source, preset, reference style, or missing dependency and rerun.

## Table Styling

No extra table beautification software is needed. The formatter uses `python-docx` to apply:

- fixed table width and column geometry
- subtle outside and inside borders
- light gray header shading
- repeated table header rows
- balanced cell padding
- centered short values and left-aligned long text
- compact but readable table fonts for wide tables

For `.txt` input, the formatter can promote simple plain-text tables into real Word tables when it sees this pattern:

```text
表 1-1 表题
序号  指标  说明
----------------
1  明火  直接触发因素
2  粉尘积聚  关键脆弱状态
```

Pipe tables in `.md` or `.txt` are also supported.
