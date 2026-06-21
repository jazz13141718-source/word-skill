# Generic Agent Instructions

Use `word-skill` as a local command-line tool for Word document conversion, formatting, audit, and render QA. It does not require a specific agent runtime.

## Minimum Workflow

1. Confirm Python 3.10+ and `python-docx` are available.
2. Run `--check-env` for the source file.
3. Run formatting or audit with `--strict` for user-facing deliverables.
4. Deliver only the final DOCX from `docx_output/` or the explicit `--final` path.
5. Keep `_docx_style_build/` files as diagnostics unless the user asks for them.

```powershell
& "<python>" "<word-skill>\scripts\format_docx_document.py" --check-env --source "<source-file>"

& "<python>" "<word-skill>\scripts\format_docx_document.py" `
  --source "<source-file>" `
  --preset report `
  --strict
```

## Template Matching

When the user asks to format one document according to a Word template, use:

```powershell
& "<python>" "<word-skill>\scripts\format_docx_document.py" `
  --source "<target-file>" `
  --style-source "<template.docx>" `
  --preset reference `
  --strict
```

## Dependency Behavior

Do not silently install Pandoc, LibreOffice, Poppler, TeX Live, fonts, or Python packages. Report missing requirements from `--check-env` and wait for user approval before installing anything.

## Agent Integration

Use `AGENT_SKILLS_DIR` or `WORD_SKILL_HOME` when installing into a non-Codex agent directory. Use `--target` when the agent has a known tools or skills folder. `agents/openai.yaml` is optional OpenAI/Codex UI metadata and is not required for generic use.
