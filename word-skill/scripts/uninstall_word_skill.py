# -*- coding: utf-8 -*-
"""Cleanly uninstall word-skill from a Codex skills directory."""
from __future__ import annotations

import argparse
import json
import os
import shutil
import sys
from pathlib import Path


SKILL_NAME = "word-skill"


def default_skills_parent() -> Path:
    codex_home = os.environ.get("CODEX_HOME")
    if codex_home:
        return Path(codex_home).expanduser() / "skills"
    return Path.home() / ".codex" / "skills"


def resolve_target(target: Path | None) -> Path:
    if target is None:
        return default_skills_parent() / SKILL_NAME
    target = target.expanduser()
    return target if target.name == SKILL_NAME else target / SKILL_NAME


def read_skill_name(path: Path) -> str | None:
    skill_md = path / "SKILL.md"
    if not skill_md.exists():
        return None
    for line in skill_md.read_text(encoding="utf-8-sig", errors="ignore").splitlines()[:20]:
        if line.strip().startswith("name:"):
            return line.split(":", 1)[1].strip().strip('"\'')
    return None


def validate_uninstall_target(target: Path):
    target = target.resolve()
    if target.name != SKILL_NAME:
        raise RuntimeError(f"Refusing to remove a directory not named {SKILL_NAME}: {target}")
    if len(target.parts) < 4:
        raise RuntimeError(f"Refusing to remove suspiciously short path: {target}")
    if read_skill_name(target) != SKILL_NAME:
        raise RuntimeError(f"Refusing to remove a directory that does not validate as {SKILL_NAME}: {target}")


def uninstall_skill(target: Path, dry_run: bool) -> dict:
    target = target.resolve()
    if not target.exists():
        return {"ok": True, "removed": False, "reason": "target does not exist", "target": str(target), "dry_run": dry_run}
    validate_uninstall_target(target)
    file_count = sum(1 for item in target.rglob("*") if item.is_file())
    result = {"ok": True, "removed": False, "target": str(target), "file_count": file_count, "dry_run": dry_run}
    if dry_run:
        result["would_remove"] = True
        return result
    shutil.rmtree(target)
    result["removed"] = True
    return result


def main() -> int:
    parser = argparse.ArgumentParser(description="Uninstall word-skill from a Codex skills directory.")
    parser.add_argument("--target", type=Path, help="Target skill directory or skills parent directory. Defaults to CODEX_HOME\\skills or ~/.codex/skills.")
    parser.add_argument("--dry-run", action="store_true", help="Print the uninstall plan without deleting files.")
    args = parser.parse_args()

    result = uninstall_skill(resolve_target(args.target), args.dry_run)
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as exc:
        print(json.dumps({"ok": False, "error": str(exc)}, ensure_ascii=False, indent=2), file=sys.stderr)
        sys.exit(1)
