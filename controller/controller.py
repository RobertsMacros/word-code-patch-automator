"""
Word VBA Patch Automator — Controller (MVP)

Runs inside Windows (e.g. Parallels VM).
Orchestrates: import VBA → run test harness → capture results → generate repair prompt.

Architecture:
  - Controller (this script) is the parent process / orchestrator.
  - word_runner.py is spawned as a child process for each test run.
  - The subprocess boundary allows real timeout enforcement:
    if Word hangs, the controller kills the child process and Word.

Dependencies: pywin32 (pip install pywin32)
"""

import json
import os
import subprocess
import sys
import time
import tempfile
from datetime import datetime
from pathlib import Path

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------

SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = SCRIPT_DIR.parent
WORD_RUNNER = SCRIPT_DIR / "word_runner.py"


def load_config():
    with open(PROJECT_ROOT / "project.json", "r", encoding="utf-8") as f:
        return json.load(f)


# ---------------------------------------------------------------------------
# Locked-path enforcement
# ---------------------------------------------------------------------------

def load_locked_paths():
    with open(PROJECT_ROOT / "locked_paths.json", "r", encoding="utf-8") as f:
        data = json.load(f)
    return data.get("locked", [])


def is_under_locked_path(rel_path, locked_paths):
    """Check if a relative path falls under any locked path."""
    rel_norm = rel_path.replace("\\", "/")
    for locked in locked_paths:
        locked_norm = locked.rstrip("/").replace("\\", "/")
        # Match exact file or anything under a locked directory
        if rel_norm == locked_norm or rel_norm.startswith(locked_norm + "/"):
            return True
        # Match exact file entries (e.g. "project.json")
        if "/" not in locked_norm and rel_norm == locked_norm:
            return True
    return False


def check_locked_paths_working_tree(locked_paths):
    """
    Check for locked-path violations in the working tree.
    Catches: modified (unstaged), staged, and untracked files under locked paths.
    Returns list of (filepath, status) tuples for violations.
    """
    violations = []

    # git status --porcelain gives us everything: staged, modified, untracked
    try:
        result = subprocess.run(
            ["git", "status", "--porcelain"],
            capture_output=True, text=True, cwd=str(PROJECT_ROOT)
        )
        if result.returncode != 0:
            return violations

        for line in result.stdout.splitlines():
            if len(line) < 4:
                continue

            status_code = line[:2]
            file_path = line[3:].strip()

            # Handle renames: "R  old -> new"
            if " -> " in file_path:
                file_path = file_path.split(" -> ", 1)[1]

            # Remove quotes if present
            file_path = file_path.strip('"')

            if is_under_locked_path(file_path, locked_paths):
                # Decode status
                if status_code[0] == "?" and status_code[1] == "?":
                    status = "untracked"
                elif status_code[0] != " ":
                    status = "staged"
                else:
                    status = "modified"
                violations.append((file_path, status))

    except Exception:
        pass

    return violations


def check_locked_paths():
    """
    Full locked-path check. Returns True if clean, False if violations found.
    Prints violations to stdout.
    """
    locked_paths = load_locked_paths()
    violations = check_locked_paths_working_tree(locked_paths)

    if violations:
        print("LOCKED PATH VIOLATION — the following locked files have changes:")
        for fpath, status in violations:
            print(f"  [{status:>9s}] {fpath}")
        print()
        print("Revert these changes before running the controller.")
        print("Aborting.")
        return False

    return True


# ---------------------------------------------------------------------------
# Fixture discovery
# ---------------------------------------------------------------------------

def discover_fixtures(fixtures_dir):
    """
    Find all fixture documents in the fixtures directory.
    Returns list of absolute paths, sorted by name.
    Supports .docm, .docx, .doc files.
    """
    fixture_path = Path(fixtures_dir)
    if not fixture_path.exists():
        return []

    valid_ext = {".docm", ".docx", ".doc"}
    fixtures = []
    for fpath in sorted(fixture_path.iterdir()):
        if fpath.suffix.lower() in valid_ext and not fpath.name.startswith("~$"):
            fixtures.append(str(fpath.resolve()))

    return fixtures


# ---------------------------------------------------------------------------
# Kill Word
# ---------------------------------------------------------------------------

def kill_word():
    """Force-kill all Word processes."""
    subprocess.run(
        ["taskkill", "/F", "/IM", "WINWORD.EXE"],
        capture_output=True
    )
    time.sleep(2)


# ---------------------------------------------------------------------------
# Test execution (subprocess with real timeout)
# ---------------------------------------------------------------------------

def run_test_pass(config, iteration, results_dir):
    """
    Run one full test pass by spawning word_runner.py as a subprocess.
    The controller enforces a real timeout — if the subprocess hangs,
    it kills the process and Word.

    Returns (all_passed: bool, report: dict).
    """
    timeout = config["controller"].get("timeout_seconds", 60)
    visible = config["controller"].get("word_visible", False)
    entry_sub = config["harness"].get("entry_sub", "RunAllTests")
    kill_on_timeout = config["controller"].get("kill_on_timeout", True)

    host_doc_path = str((PROJECT_ROOT / config["host_doc"]).resolve())
    src_dir = str((PROJECT_ROOT / config.get("src_dir", "src")).resolve())
    harness_dir = str((PROJECT_ROOT / "harness").resolve())
    fixtures_dir = str((PROJECT_ROOT / config.get("fixtures_dir", "tests/fixtures")).resolve())

    result_file = str(results_dir / f"harness_output_{iteration:03d}.json")
    runner_output = str(results_dir / f"runner_output_{iteration:03d}.json")

    # Discover fixture documents
    fixtures = discover_fixtures(fixtures_dir)

    # Build the run config for word_runner.py
    run_config = {
        "host_doc_path": host_doc_path,
        "src_dir": src_dir,
        "harness_dir": harness_dir,
        "entry_sub": entry_sub,
        "word_visible": visible,
        "result_file": result_file,
        "runner_output_file": runner_output,
        "fixtures": fixtures
    }

    # Write run config to a temp file
    run_config_path = str(results_dir / f"_run_config_{iteration:03d}.json")
    with open(run_config_path, "w", encoding="utf-8") as f:
        json.dump(run_config, f, indent=2)

    # Spawn the worker subprocess
    start = time.time()
    word_killed = False
    timed_out = False
    error_msg = None
    runner_result = None

    try:
        proc = subprocess.Popen(
            [sys.executable, str(WORD_RUNNER), run_config_path],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE
        )

        try:
            stdout, stderr = proc.communicate(timeout=timeout)
        except subprocess.TimeoutExpired:
            timed_out = True
            proc.kill()
            proc.wait(timeout=5)

            if kill_on_timeout:
                kill_word()
                word_killed = True

            error_msg = f"TIMEOUT: test run exceeded {timeout}s limit"
            if word_killed:
                error_msg += " (Word killed)"

    except Exception as e:
        error_msg = f"Failed to spawn word_runner: {e}"

    elapsed = time.time() - start

    # Read runner output if it exists
    if os.path.exists(runner_output):
        try:
            with open(runner_output, "r", encoding="utf-8") as f:
                runner_result = json.load(f)
        except Exception as e:
            error_msg = error_msg or f"Failed to read runner output: {e}"

    # Build the report
    if runner_result and runner_result.get("error"):
        error_msg = error_msg or runner_result["error"]

    tests = []
    fixture_results = []
    if runner_result:
        tests = runner_result.get("tests", [])
        fixture_results = runner_result.get("fixtures", [])

    passed = sum(1 for t in tests if t.get("passed"))
    failed = sum(1 for t in tests if not t.get("passed"))
    total = len(tests)
    all_passed = (failed == 0 and total > 0 and error_msg is None
                  and not timed_out)

    report = {
        "iteration": iteration,
        "timestamp": datetime.now().isoformat(),
        "elapsed_seconds": round(elapsed, 3),
        "error": error_msg,
        "timed_out": timed_out,
        "word_killed": word_killed,
        "fixtures_tested": [os.path.basename(f) for f in fixtures],
        "fixture_results": fixture_results,
        "summary": {
            "total": total,
            "passed": passed,
            "failed": failed,
            "all_passed": all_passed
        },
        "tests": tests,
        "raw_runner_output": runner_result
    }

    return all_passed, report


# ---------------------------------------------------------------------------
# Reporting
# ---------------------------------------------------------------------------

def generate_summary(report):
    """Generate a concise human-readable summary string."""
    lines = []
    iteration = report["iteration"]
    elapsed = report["elapsed_seconds"]

    lines.append(f"=== Test Run Summary (iteration {iteration}) ===")
    lines.append(f"Time:    {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    lines.append(f"Elapsed: {elapsed:.2f}s")

    if report.get("timed_out"):
        lines.append(f"TIMEOUT: run exceeded limit")
    if report.get("word_killed"):
        lines.append(f"Word was force-killed")
    if report.get("error"):
        lines.append(f"ERROR: {report['error']}")

    # Fixture info
    fixtures = report.get("fixtures_tested", [])
    if fixtures:
        lines.append(f"Fixtures: {', '.join(fixtures)}")

    s = report["summary"]
    lines.append(f"Tests: {s['total']} total, {s['passed']} passed, {s['failed']} failed")

    # Failed tests
    failed_tests = [t for t in report.get("tests", []) if not t.get("passed")]
    if failed_tests:
        lines.append("")
        lines.append("Failed tests:")
        for t in failed_tests:
            name = t.get("name", "?")
            msg = t.get("message", "")
            dur = t.get("duration_ms", "?")
            fixture = t.get("fixture", "")
            prefix = f"  FAIL: {name}"
            if fixture:
                prefix += f" [{fixture}]"
            prefix += f" ({dur}ms)"
            if msg:
                prefix += f" — {msg}"
            lines.append(prefix)

    # Timed-out individual tests (VBA-level)
    timed_out_tests = [t for t in report.get("tests", []) if t.get("timed_out")]
    if timed_out_tests:
        lines.append("")
        lines.append("Timed-out tests:")
        for t in timed_out_tests:
            lines.append(f"  TIMEOUT: {t.get('name', '?')}")

    lines.append("")
    lines.append(f"Result: {'PASS' if s['all_passed'] else 'FAIL'}")
    lines.append("=" * 48)

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Repair prompt generation
# ---------------------------------------------------------------------------

def generate_repair_prompt(report, config, iteration):
    """Generate a repair prompt file for Claude Code CLI."""
    template_path = PROJECT_ROOT / "controller" / "templates" / "repair_prompt.md"
    if template_path.exists():
        with open(template_path, "r", encoding="utf-8") as f:
            template = f.read()
    else:
        template = DEFAULT_REPAIR_TEMPLATE

    failed_tests = [t for t in report.get("tests", []) if not t.get("passed")]
    failed_detail = json.dumps(failed_tests, indent=2)

    src_dir = config.get("src_dir", "src")
    mutable = config.get("mutable_paths", ["src/"])

    prompt = template.format(
        iteration=iteration,
        max_iterations=config["controller"]["max_iterations"],
        failed_count=len(failed_tests),
        total_count=report["summary"]["total"],
        failed_tests_json=failed_detail,
        error_message=report.get("error") or "None",
        src_dir=src_dir,
        mutable_paths="\n".join(f"  - {p}" for p in mutable),
        elapsed=report["elapsed_seconds"],
        timed_out="Yes" if report.get("timed_out") else "No",
        word_killed="Yes" if report.get("word_killed") else "No",
        full_report_json=json.dumps(report, indent=2)
    )

    return prompt


DEFAULT_REPAIR_TEMPLATE = """# VBA Repair Prompt — Iteration {iteration}/{max_iterations}

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
"""


# ---------------------------------------------------------------------------
# Main controller loop
# ---------------------------------------------------------------------------

def main():
    config = load_config()
    max_iters = config["controller"]["max_iterations"]
    timeout = config["controller"]["timeout_seconds"]
    results_dir = PROJECT_ROOT / config.get("results_dir", "results")
    results_dir.mkdir(parents=True, exist_ok=True)

    # Verify host document exists
    host_doc = PROJECT_ROOT / config["host_doc"]
    if not host_doc.exists():
        print(f"Error: Host document not found: {host_doc}")
        print(f"Create a macro-enabled document at: {host_doc}")
        sys.exit(1)

    # Verify fixtures directory
    fixtures_dir = PROJECT_ROOT / config.get("fixtures_dir", "tests/fixtures")
    fixtures = discover_fixtures(str(fixtures_dir))

    print("=" * 60)
    print("Word VBA Patch Automator — Controller (MVP)")
    print(f"Project:    {config['project_name']}")
    print(f"Host doc:   {host_doc}")
    print(f"Fixtures:   {len(fixtures)} found in {fixtures_dir}")
    print(f"Timeout:    {timeout}s")
    print(f"Max iters:  {max_iters}")
    print("=" * 60)

    if not fixtures:
        print(f"\nWarning: No fixture documents found in {fixtures_dir}")
        print("Tests that require fixtures will be skipped.")

    # Check locked paths BEFORE running
    if not check_locked_paths():
        sys.exit(1)

    run_reports = []

    for iteration in range(1, max_iters + 1):
        print(f"\n--- Iteration {iteration}/{max_iters} ---")

        all_passed, report = run_test_pass(config, iteration, results_dir)
        run_reports.append(report)

        # Write detailed JSON report
        report_file = results_dir / f"report_{iteration:03d}.json"
        with open(report_file, "w", encoding="utf-8") as f:
            json.dump(report, f, indent=2)

        # Write + print human summary
        summary = generate_summary(report)
        summary_file = results_dir / f"summary_{iteration:03d}.txt"
        with open(summary_file, "w", encoding="utf-8") as f:
            f.write(summary)
        print(summary)

        if all_passed:
            print(f"\nAll tests passed on iteration {iteration}. Done!")
            break

        if iteration < max_iters:
            # Generate repair prompt
            prompt = generate_repair_prompt(report, config, iteration)
            prompt_file = results_dir / f"repair_prompt_{iteration:03d}.md"
            with open(prompt_file, "w", encoding="utf-8") as f:
                f.write(prompt)
            print(f"\nRepair prompt: {prompt_file}")
            print("[Semi-automatic] Run Claude Code CLI, then re-run the controller:")
            print(f"  claude -p {prompt_file}")
            break
        else:
            print(f"\nMax iterations ({max_iters}) reached. Tests still failing.")

    # Write combined run log
    combined_file = results_dir / "run_log.json"
    with open(combined_file, "w", encoding="utf-8") as f:
        json.dump({
            "project": config["project_name"],
            "timestamp": datetime.now().isoformat(),
            "iterations": run_reports
        }, f, indent=2)

    print(f"\nRun log: {combined_file}")


if __name__ == "__main__":
    main()
