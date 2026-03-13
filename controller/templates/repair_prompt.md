# VBA Repair Prompt — Iteration {iteration}/{max_iterations}

## Status
- **{failed_count}** of **{total_count}** tests failed
- Elapsed: {elapsed}s
- Timed out: {timed_out}
- Word killed: {word_killed}
- Error: {error_message}

## Rules
1. You may ONLY modify files under: `{src_dir}/`
2. Allowed mutable paths:
{mutable_paths}
3. Do NOT modify any test files, fixtures, expected outputs, harness code, or controller code.
4. Do NOT add new files outside the mutable paths.
5. Focus on fixing the failing tests below.
6. Make minimal, targeted fixes — do not refactor unrelated code.

## Failed Tests

```json
{failed_tests_json}
```

## Full Test Report

```json
{full_report_json}
```

## Instructions
1. Read the failing test details above carefully.
2. Read the relevant VBA source files under `{src_dir}/`.
3. Identify the root cause of each failure.
4. Make minimal, targeted fixes in the source files.
5. Do not change any files outside `{src_dir}/`.
6. Save all modified files when done.
