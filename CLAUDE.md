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

<!-- project-coordination:reporting-v2:start -->
## Shared progress reporting

This repository uses `project-coordination/v1`. At the start of substantive work,
read `.coordination/README.md`, existing task breadcrumbs and the shared claim.
Before changing code, run `.coordination/report.py start` with the exact task UUID,
a unique actor UUID for this chat/run, title and scope. Start only after it exits
successfully and confirms the claim on GitHub. Another active actor means wait;
never duplicate, steal or silently expire a claim. Preserve IDs through handover.

Use the reporter for substantive progress, blockers and outcomes, including local
and unpushed work. Report at take-on, material changes, before handover/ending and
when the outcome changes; do not publish polling noise. Only request Robert's
input for a specific decision, approval, access, file, token or physical action.
Agent review, testing and ordinary implementation are agent work.

Publishing sanitised coordination metadata to the private coordination repository
is authorised separately from publishing project code. Never push implementation,
secrets, personal/case/financial/medical data or private local paths through this
reporting rule. Done assertions require exact-scope evidence; commit, delivery and
process exit alone do not prove completion. Read the reporter receipt; on failure,
retain the local breadcrumb and report publication uncertainty. Never claim sync.
<!-- project-coordination:reporting-v2:end -->
