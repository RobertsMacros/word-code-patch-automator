# CLAUDE.md — Instructions for Claude Code

## Project
Word VBA Patch Automator — a local automation loop for testing and repairing VBA macro code.

## Immutable Paths (DO NOT MODIFY)
The following paths are locked. You must NEVER create, edit, delete, or overwrite files in these locations:

- `tests/` — all test definitions, fixtures, and expected outputs
- `harness/` — the VBA test harness modules
- `controller/` — the Python controller and templates
- `project.json` — project configuration
- `locked_paths.json` — locked path definitions
- `CLAUDE.md` — this file

## Mutable Paths (YOU MAY MODIFY)
You may only modify files under:

- `src/` — VBA source modules (.bas, .cls, .frm)

## Repair Task Rules
When given a repair prompt:
1. Read the failing test details carefully
2. Read the relevant source files under `src/`
3. Make minimal, targeted fixes to address the failures
4. Do NOT refactor, rename, or reorganise code beyond what is needed
5. Do NOT add new files outside `src/`
6. Do NOT modify tests, fixtures, expected outputs, harness, or controller code
