# Agent Guide

`word-skill` is usable by any agent that can run local commands. The core workflow is a Python CLI; the OpenAI/Codex metadata in `word-skill/agents/openai.yaml` is optional and not required for other agents.

## How To Use

1. Locate the checked-out repository or installed `word-skill` directory.
2. Run the environment check before formatting:

```powershell
& "<python>" ".\word-skill\scripts\format_docx_document.py" --check-env --source ".\input.md"
```

3. Format with a preset:

```powershell
& "<python>" ".\word-skill\scripts\format_docx_document.py" `
  --source ".\input.md" `
  --preset report `
  --strict
```

4. For template-style formatting, pass a reference Word file:

```powershell
& "<python>" ".\word-skill\scripts\format_docx_document.py" `
  --source ".\target.docx" `
  --style-source ".\template.docx" `
  --preset reference `
  --strict
```

5. Deliver only the final DOCX from `docx_output/` or the explicit `--final` path. Treat `_docx_style_build/` as diagnostics.

## Install Locations

No universal agent skill directory exists. Prefer an explicit target:

```powershell
& "<python>" ".\word-skill\scripts\install_word_skill.py" --target "D:\SomeAgent\skills"
```

The installer also accepts `AGENT_SKILLS_DIR` or `WORD_SKILL_HOME` for generic agent setups. If neither is set, it falls back to Codex-compatible defaults for backward compatibility.

## Quality Gate

Use `--strict` for user-facing deliverables. Before saying the document is ready, inspect `run_report.json` and confirm `quality_audit.ok` is `true`. For important documents, inspect rendered PNG pages, especially around wide tables, images, references, and dense layouts.
