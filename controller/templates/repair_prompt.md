# VBA Repair Prompt — Iteration {iteration}/{max_iterations}

## Status
- **{failed_count}** of **{total_count}** tests failed
- Elapsed: {elapsed}s
- Timed out: {timed_out}
- Word killed: {word_killed}
- Error: {error_message}

## Rules
You may modify files under these paths:
{mutable_paths}

You must NOT modify these protected paths:
{locked_paths}

Make minimal, targeted fixes. Do not refactor unrelated code.

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
2. Read the relevant source files.
3. Identify the root cause of each failure.
4. Make minimal, targeted fixes.
5. Save all modified files when done.
