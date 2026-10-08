# CLAUDE.md — Instructions for Claude Code

## Project
Word VBA Patch Automator — a local automation loop for testing and repairing VBA macro code.

## Protected Paths (DO NOT MODIFY)
The following paths are protected. You must NEVER create, edit, delete, or overwrite files in these locations:

- `tests/fixtures/` — fixture documents used as test inputs
- `tests/expected/` — expected outputs for regression tests
- `project.json` — project configuration
- `CLAUDE.md` — this file

## Mutable Paths (YOU MAY MODIFY)
You may modify files under:

- `src/` — VBA source modules (.bas, .cls, .frm)
- `controller/` — Python controller and runner scripts
- `harness/` — VBA test harness modules

## Repair Task Rules
When given a repair prompt:
1. Read the failing test details carefully
2. Read the relevant source files
3. Make minimal, targeted fixes to address the failures
4. Do NOT refactor, rename, or reorganise code beyond what is needed
5. Do NOT add new files outside the mutable paths
6. Do NOT modify fixture documents or expected outputs

<!-- project-coordination:v1:start -->
## Project coordination

Read [.coordination/README.md](.coordination/README.md) at the start of substantive work. Leave immutable JSON breadcrumbs under `.coordination/events/` for material progress, blockers and verified completion. Reuse the exact incoming task/request and roadmap milestone IDs across chats. Write `done` only with evidence for the entire recorded scope; delivery, partial work and unverified deployment/device checks are not completion. Preserve other actors' records, existing project rules, consent, privacy and publication limits. This applies to Codex, Claude and other project agents; no Page connection or chat-to-chat message is needed.
<!-- project-coordination:v1:end -->
