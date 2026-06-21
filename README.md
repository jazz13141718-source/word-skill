# word-skill

`word-skill` is a Codex skill for turning AI-generated text, Markdown drafts, and existing office documents into polished Word deliverables. It supports Chinese and English content, multiple document presets, reference-DOCX style matching, separated final/build outputs, and render-based quality checks.

`word-skill` 是一个面向 Codex/Agent 工作流的 Word 文档格式化 skill。它适合把 AI 对话产出的文本、Markdown 草稿、已有 Word 文档、报告、汇报材料、发言稿、论文等内容整理成可交付的 `.docx` 文件。

> License note: this repository is public source / source-available under a non-commercial license. Commercial use requires separate written permission. See [LICENSE](LICENSE).

## Highlights

- Convert Markdown and plain text to DOCX through Pandoc.
- Reformat existing DOCX files without mixing final files and temporary build artifacts.
- Convert DOC, RTF, and ODT through LibreOffice when available.
- Match the layout and styles of another Word document with `--style-source`.
- Use built-in presets for papers, reports, briefings, speeches, general documents, and English documents.
- Save custom presets for repeatable document styles.
- Normalize body styles, headings, page setup, footers, tables, captions, and inline citation superscripts.
- Audit output with machine-readable `quality_audit` results.
- Render pages through LibreOffice and `pdftoppm` so agents can visually inspect final layout.

## Suitable For

- Exporting AI conversation content to a clean Word document.
- Formatting academic papers, reports, briefings, proposals, meeting materials, and speeches.
- Reformatting one Word file according to another Word file's style.
- Checking whether a generated DOCX has broken tables, missing images, wrong body styles, blank renders, or citation-format issues.
- Packaging a repeatable Word-formatting workflow for another agent.

## Supported Inputs

| Input | Support | Notes |
| --- | --- | --- |
| `.md`, `.markdown`, `.mdown` | Yes | Uses Pandoc before post-processing. |
| `.txt` | Yes | Uses Pandoc and applies the selected preset. |
| `.docx` | Yes | Can reformat or audit directly. |
| `.doc`, `.rtf`, `.odt` | Yes | Requires LibreOffice for conversion. |

## Requirements

Required:

- Python 3.10 or newer.
- Python packages listed in [requirements.txt](requirements.txt).
- Pandoc for Markdown and text input.
- LibreOffice for DOC/RTF/ODT conversion and render QA.
- `pdftoppm` for PNG page rendering. It can come from Poppler or TeX Live.

Recommended:

- Chinese fonts such as SimSun, SimHei, Microsoft YaHei, and Times New Roman for Chinese presets.
- Times New Roman or Arial for English presets.
- Pillow for stronger blank-page / low-ink render checks.

The skill does not silently install system tools. Run `--check-env` first, review missing dependencies, and install them only after user approval.

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

Uninstall cleanly:

```powershell
& "<python>" ".\word-skill\scripts\uninstall_word_skill.py" --dry-run
& "<python>" ".\word-skill\scripts\uninstall_word_skill.py"
```

## Quick Start

Check the environment:

```powershell
& "<python>" ".\word-skill\scripts\format_docx_document.py" --check-env --source ".\paper.md"
```

Format a Chinese report:

```powershell
& "<python>" ".\word-skill\scripts\format_docx_document.py" `
  --source ".\paper.md" `
  --preset report `
  --strict
```

Format an English report:

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

Format one Word file according to another Word file:

```powershell
& "<python>" ".\word-skill\scripts\format_docx_document.py" `
  --source ".\target.docx" `
  --style-source ".\reference.docx" `
  --preset reference `
  --strict
```

Save and reuse a custom preset:

```powershell
& "<python>" ".\word-skill\scripts\format_docx_document.py" `
  --source ".\draft.md" `
  --preset report `
  --preset-file ".\my-format.json" `
  --save-preset-name "my-report-style" `
  --strict
```

## Built-In Presets

| Preset | Use Case |
| --- | --- |
| `auto` | Infer a suitable preset from source and options. |
| `general` | General AI-exported documents and mixed materials. |
| `academic-paper` | Papers and thesis-like manuscripts. |
| `report` | Formal reports and longer written materials. |
| `briefing` | Concise briefings and executive summaries. |
| `speech` | Speeches, remarks, and speaking scripts. |
| `english-general` | General English documents. |
| `english-academic` | English academic manuscripts. |
| `english-report` | English reports and formal documents. |
| `english-briefing` | English briefings and summaries. |
| `english-speech` | English speeches and scripts. |
| `reference` | Match styles from `--style-source`. |

## Output Policy

- Final DOCX files go to `docx_output/` by default, or to `--output-dir` / `--final`.
- Chinese presets default to `<source>_格式化.docx`.
- English presets default to `<source>_formatted.docx`.
- Build artifacts go to `_docx_style_build/<run-id>/`.
- Deliver only the final DOCX unless diagnostics are explicitly requested.
- For important deliverables, inspect rendered PNG pages before delivery.

## Quality Gate

The formatter writes a `run_report.json` file under the build directory. The key field is `quality_audit`:

```json
{
  "quality_audit": {
    "ok": true,
    "findings": [],
    "warnings": [],
    "metrics": {}
  }
}
```

Use `--strict` for user-facing deliverables. If `quality_audit.ok` is false, inspect the findings, repair the document or preset, and rerun before delivery.

The audit checks common delivery risks, including:

- Body text accidentally using the wrong Word style.
- Missing images or broken external image references.
- Malformed or overly wide tables.
- Missing render outputs.
- Blank or nearly blank rendered pages.
- Inline numeric citations that were not superscripted.
- Final output accidentally placed inside a build directory.

## Repository Structure

```text
word-skill-repo/
  word-skill/
    SKILL.md
    agents/openai.yaml
    assets/reference.docx
    references/environment-and-output.md
    scripts/format_docx_document.py
    scripts/install_word_skill.py
    scripts/uninstall_word_skill.py
  tests/test_word_skill_smoke.py
  requirements.txt
  LICENSE
  README.md
```

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

`word-skill Non-Commercial License 1.0`. See [LICENSE](LICENSE).

This is not an OSI-approved open source license because it restricts commercial use. It is intended for public, non-commercial collaboration.
