# -*- coding: utf-8 -*-
"""Install word-skill into an agent skills or tools directory."""
from __future__ import annotations

import argparse
import json
import os
import shutil
import sys
from pathlib import Path


SKILL_NAME = "word-skill"
SKIP_DIRS = {"__pycache__", ".pytest_cache", "_docx_style_build", "docx_output"}
SKIP_SUFFIXES = {".pyc", ".pyo", ".tmp", ".bak"}


def as_skill_target(path: Path) -> Path:
    path = path.expanduser()
    return path if path.name == SKILL_NAME else path / SKILL_NAME


def default_target() -> Path:
    word_skill_home = os.environ.get("WORD_SKILL_HOME")
    if word_skill_home:
        return as_skill_target(Path(word_skill_home))
    agent_skills_dir = os.environ.get("AGENT_SKILLS_DIR")
    if agent_skills_dir:
        return as_skill_target(Path(agent_skills_dir))
    codex_home = os.environ.get("CODEX_HOME")
    if codex_home:
        return Path(codex_home).expanduser() / "skills" / SKILL_NAME
    return Path.home() / ".codex" / "skills" / SKILL_NAME


def resolve_target(target: Path | None) -> Path:
    if target is None:
        return default_target()
    return as_skill_target(target)


def read_skill_name(path: Path) -> str | None:
    skill_md = path / "SKILL.md"
    if not skill_md.exists():
        return None
    for line in skill_md.read_text(encoding="utf-8-sig", errors="ignore").splitlines()[:20]:
        if line.strip().startswith("name:"):
            return line.split(":", 1)[1].strip().strip('"\'')
    return None


def is_word_skill_dir(path: Path) -> bool:
    return path.is_dir() and read_skill_name(path) == SKILL_NAME


def ignore_names(_directory: str, names: list[str]) -> set[str]:
    ignored: set[str] = set()
    for name in names:
        if name in SKIP_DIRS or Path(name).suffix.lower() in SKIP_SUFFIXES:
            ignored.add(name)
    return ignored


def count_files(path: Path) -> int:
    return sum(1 for item in path.rglob("*") if item.is_file())


def install_skill(source: Path, target: Path, dry_run: bool) -> dict:
    source = source.resolve()
    target = target.resolve()
    if not is_word_skill_dir(source):
        raise RuntimeError(f"Source is not a valid {SKILL_NAME} skill: {source}")
    if target.exists() and not is_word_skill_dir(target):
        raise RuntimeError(f"Refusing to replace a non-{SKILL_NAME} directory: {target}")
    if source == target:
        return {"ok": True, "installed": False, "reason": "source already equals target", "source": str(source), "target": str(target)}

    parent = target.parent
    staging = parent / f".{SKILL_NAME}-install-{os.getpid()}"
    plan = {
        "ok": True,
        "dry_run": dry_run,
        "source": str(source),
        "target": str(target),
        "staging": str(staging),
        "will_replace_existing": target.exists(),
        "file_count": count_files(source),
    }
    if dry_run:
        return plan

    parent.mkdir(parents=True, exist_ok=True)
    if staging.exists():
        shutil.rmtree(staging)
    shutil.copytree(source, staging, ignore=ignore_names)
    if read_skill_name(staging) != SKILL_NAME:
        shutil.rmtree(staging)
        raise RuntimeError("Staged copy failed validation.")
    if target.exists():
        shutil.rmtree(target)
    staging.rename(target)
    plan["installed"] = True
    plan["installed_file_count"] = count_files(target)
    return plan


def main() -> int:
    parser = argparse.ArgumentParser(description="Install the word-skill package into an agent skills or tools directory.")
    parser.add_argument("--source", type=Path, default=Path(__file__).resolve().parent.parent, help="Source skill directory. Defaults to this script's parent skill.")
    parser.add_argument("--target", type=Path, help="Target skill directory or skills parent directory. Defaults to WORD_SKILL_HOME, AGENT_SKILLS_DIR, CODEX_HOME\\skills\\word-skill, or ~/.codex/skills/word-skill.")
    parser.add_argument("--dry-run", action="store_true", help="Print the install plan without copying files.")
    args = parser.parse_args()

    result = install_skill(args.source, resolve_target(args.target), args.dry_run)
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as exc:
        print(json.dumps({"ok": False, "error": str(exc)}, ensure_ascii=False, indent=2), file=sys.stderr)
        sys.exit(1)
