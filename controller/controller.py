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

def is_under_locked_path(rel_path, locked_paths):
    """Check if a relative path falls under any locked path."""
    rel_norm = rel_path.replace("\\", "/")
    for locked in locked_paths:
        locked_norm = locked.rstrip("/").replace("\\", "/")
        if rel_norm == locked_norm or rel_norm.startswith(locked_norm + "/"):
            return True
    return False


def check_locked_paths_working_tree(locked_paths):
    """
    Check for locked-path violations in the working tree.
    Catches: modified (unstaged), staged, and untracked files under locked paths.
    Returns list of (filepath, status) tuples for violations.
    """
    violations = []
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

            file_path = file_path.strip('"')

            if is_under_locked_path(file_path, locked_paths):
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


def check_locked_paths(config):
    """
    Full locked-path check. Returns True if clean, False if violations found.
    """
    locked_paths = config.get("locked_paths", [])
    if not locked_paths:
        return True

    violations = check_locked_paths_working_tree(locked_paths)

    if violations:
        print("LOCKED PATH VIOLATION — the following protected files have changes:")
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
    harness_dir = str((PROJECT_ROOT / config.get("harness_dir", "harness")).resolve())
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

    if runner_result and runner_result.get("error"):
        error_msg = error_msg or runner_result["error"]

    # Build fixture-centric report
    report = build_report(
        config, iteration, elapsed, fixtures,
        runner_result, error_msg, timed_out, word_killed
    )

    return report["summary"]["all_passed"], report


def build_report(config, iteration, elapsed, fixture_paths,
                 runner_result, error_msg, timed_out, word_killed):
    """
    Build the canonical report structure.
    Fixture-centric: each fixture gets a top-level entry with
    name, status, elapsed, and errors.
    """
    fixture_results = []
    tests = []

    if runner_result:
        tests = runner_result.get("tests", [])
        raw_fixtures = runner_result.get("fixtures", [])

        # Reshape raw fixture results into the canonical format
        for rf in raw_fixtures:
            name = rf.get("fixture", rf.get("name", "unknown"))
            failed = rf.get("failed", 0)
            fx_elapsed = rf.get("elapsed_seconds", 0)
            fx_error = rf.get("error", "")

            # Collect error messages for this fixture from the test list
            errors = []
            if fx_error:
                errors.append(fx_error)
            for t in tests:
                if t.get("fixture") == name and not t.get("passed"):
                    errors.append(f"{t.get('name', '?')}: {t.get('message', '')}")

            fixture_results.append({
                "name": name,
                "status": "fail" if (failed > 0 or fx_error) else "pass",
                "elapsed": fx_elapsed,
                "passed": rf.get("passed", 0),
                "failed": failed,
                "errors": errors if errors else []
            })

    # Summarise fixtures
    fx_total = len(fixture_results)
    fx_passed = sum(1 for f in fixture_results if f["status"] == "pass")
    fx_failed = sum(1 for f in fixture_results if f["status"] == "fail")

    # Summarise all tests (including non-fixture smoke/logic tests)
    test_total = len(tests)
    test_passed = sum(1 for t in tests if t.get("passed"))
    test_failed = sum(1 for t in tests if not t.get("passed"))

    all_passed = (test_failed == 0 and test_total > 0
                  and error_msg is None and not timed_out)

    return {
        "project": config["project_name"],
        "iteration": iteration,
        "timestamp": datetime.now().isoformat(),
        "elapsed_seconds": round(elapsed, 3),
        "error": error_msg,
        "timed_out": timed_out,
        "word_killed": word_killed,
        "fixtures": fixture_results,
        "tests": tests,
        "summary": {
            "fixtures_total": fx_total,
            "fixtures_passed": fx_passed,
            "fixtures_failed": fx_failed,
            "tests_total": test_total,
            "tests_passed": test_passed,
            "tests_failed": test_failed,
            "timed_out": 1 if timed_out else 0,
            "all_passed": all_passed
        }
    }


# ---------------------------------------------------------------------------
# Reporting
# ---------------------------------------------------------------------------

def generate_summary(report):
    """Generate a concise human-readable summary."""
    lines = []
    s = report["summary"]
    elapsed = report["elapsed_seconds"]

    lines.append("Run summary")
    lines.append("-----------")
    lines.append(f"iteration: {report['iteration']}")
    lines.append(f"fixtures:  {s['fixtures_total']}")
    lines.append(f"tests:     {s['tests_total']}")
    lines.append(f"passed:    {s['tests_passed']}")
    lines.append(f"failed:    {s['tests_failed']}")
    lines.append(f"timeouts:  {s['timed_out']}")
    lines.append(f"elapsed:   {elapsed:.1f}s")

    if report.get("word_killed"):
        lines.append("word:      killed")

    if report.get("error"):
        lines.append(f"error:     {report['error']}")

    # Failed fixtures
    failed_fixtures = [f for f in report.get("fixtures", []) if f["status"] == "fail"]
    if failed_fixtures:
        lines.append("")
        lines.append("Failed fixtures:")
        for f in failed_fixtures:
            lines.append(f"  - {f['name']}")

    # Failed tests (including non-fixture)
    failed_tests = [t for t in report.get("tests", []) if not t.get("passed")]
    if failed_tests:
        lines.append("")
        lines.append("Failed tests:")
        for t in failed_tests:
            name = t.get("name", "?")
            msg = t.get("message", "")
            fixture = t.get("fixture", "")
            line = f"  - {name}"
            if fixture:
                line += f" [{fixture}]"
            if msg:
                line += f" — {msg}"
            lines.append(line)

    lines.append("")
    lines.append(f"result: {'PASS' if s['all_passed'] else 'FAIL'}")
    lines.append(f"report: results/report_{report['iteration']:03d}.json")

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

    mutable = config.get("mutable_paths", ["src/"])
    locked = config.get("locked_paths", [])

    prompt = template.format(
        iteration=iteration,
        max_iterations=config["controller"]["max_iterations"],
        failed_count=len(failed_tests),
        total_count=report["summary"]["tests_total"],
        failed_tests_json=json.dumps(failed_tests, indent=2),
        error_message=report.get("error") or "None",
        mutable_paths="\n".join(f"  - `{p}`" for p in mutable),
        locked_paths="\n".join(f"  - `{p}`" for p in locked),
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

    print("=" * 50)
    print("Word VBA Patch Automator")
    print(f"Project:   {config['project_name']}")
    print(f"Host doc:  {host_doc}")
    print(f"Fixtures:  {len(fixtures)} in {fixtures_dir}")
    print(f"Timeout:   {timeout}s")
    print(f"Max iters: {max_iters}")
    print("=" * 50)

    if not fixtures:
        print(f"\nWarning: No fixture documents found in {fixtures_dir}")
        print("Fixture-based tests will be skipped.")

    # Check locked paths BEFORE running
    if not check_locked_paths(config):
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

    # Regenerate the whole-project code export
    refresh_code_export()


def refresh_code_export():
    """Regenerate project_code_export.txt so it stays current with every patch."""
    export_script = PROJECT_ROOT / "export_all_code.py"
    if export_script.exists():
        try:
            subprocess.run(
                [sys.executable, str(export_script)],
                cwd=str(PROJECT_ROOT),
                capture_output=True, timeout=10
            )
        except Exception:
            pass


if __name__ == "__main__":
    main()
