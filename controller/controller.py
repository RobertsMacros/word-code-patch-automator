"""
Word VBA Patch Automator — Controller (MVP)

Runs inside Windows (e.g. Parallels VM).
Orchestrates: import VBA → run test harness → capture results → generate repair prompt.

Dependencies: pywin32 (pip install pywin32)
"""

import json
import os
import subprocess
import sys
import time
import shutil
import signal
from datetime import datetime
from pathlib import Path

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = SCRIPT_DIR.parent

def load_config():
    cfg_path = PROJECT_ROOT / "project.json"
    with open(cfg_path, "r", encoding="utf-8") as f:
        return json.load(f)

# ---------------------------------------------------------------------------
# Locked-path enforcement
# ---------------------------------------------------------------------------

def load_locked_paths():
    lp_path = PROJECT_ROOT / "locked_paths.json"
    with open(lp_path, "r", encoding="utf-8") as f:
        data = json.load(f)
    return data.get("locked", [])


def check_locked_paths(changed_files, locked_paths):
    """Return list of violations: changed files that fall under locked paths."""
    violations = []
    for fpath in changed_files:
        rel = os.path.relpath(fpath, PROJECT_ROOT).replace("\\", "/")
        for locked in locked_paths:
            locked_norm = locked.replace("\\", "/")
            if rel == locked_norm or rel.startswith(locked_norm):
                violations.append(rel)
                break
    return violations


def get_changed_files_since(commit_hash):
    """Return list of files changed since the given commit (absolute paths)."""
    try:
        result = subprocess.run(
            ["git", "diff", "--name-only", commit_hash, "HEAD"],
            capture_output=True, text=True, cwd=str(PROJECT_ROOT)
        )
        if result.returncode != 0:
            return []
        files = [
            str(PROJECT_ROOT / line.strip())
            for line in result.stdout.strip().splitlines()
            if line.strip()
        ]
        return files
    except Exception:
        return []


def get_current_commit():
    try:
        result = subprocess.run(
            ["git", "rev-parse", "HEAD"],
            capture_output=True, text=True, cwd=str(PROJECT_ROOT)
        )
        return result.stdout.strip() if result.returncode == 0 else None
    except Exception:
        return None


# ---------------------------------------------------------------------------
# Word COM automation
# ---------------------------------------------------------------------------

def get_word_app(visible=False):
    """Launch or connect to Word via COM."""
    import win32com.client
    try:
        word = win32com.client.GetActiveObject("Word.Application")
    except Exception:
        word = win32com.client.Dispatch("Word.Application")
    word.Visible = visible
    word.DisplayAlerts = 0  # wdAlertsNone
    return word


def kill_word():
    """Force-kill all Word processes."""
    subprocess.run(
        ["taskkill", "/F", "/IM", "WINWORD.EXE"],
        capture_output=True
    )
    time.sleep(2)


def import_vba_modules(word, doc, src_dir):
    """
    Import all .bas / .cls / .frm files from src_dir into the document's
    VBA project, replacing existing modules of the same name.
    """
    vb_project = doc.VBProject
    src_path = Path(src_dir)

    # Map of extension → VBA component type constants
    # 1=vbext_ct_StdModule, 2=vbext_ct_ClassModule, 3=vbext_ct_MSForm
    ext_types = {".bas": 1, ".cls": 2, ".frm": 3}

    imported = []
    for fpath in sorted(src_path.iterdir()):
        ext = fpath.suffix.lower()
        if ext not in ext_types:
            continue

        mod_name = fpath.stem

        # Remove existing module if present (skip built-in document modules)
        try:
            existing = vb_project.VBComponents(mod_name)
            comp_type = existing.Type
            # 100 = vbext_ct_Document — cannot remove these
            if comp_type != 100:
                vb_project.VBComponents.Remove(existing)
        except Exception:
            pass

        # Import
        vb_project.VBComponents.Import(str(fpath.resolve()))
        imported.append(fpath.name)

    return imported


def import_harness_modules(word, doc, harness_dir):
    """Import test harness modules from the harness/ directory."""
    return import_vba_modules(word, doc, harness_dir)


def export_vba_modules(doc, dest_dir):
    """
    Export all VBA components from the document to dest_dir.
    Skips document modules (ThisDocument, Sheet objects).
    """
    dest = Path(dest_dir)
    dest.mkdir(parents=True, exist_ok=True)

    ext_map = {1: ".bas", 2: ".cls", 3: ".frm"}
    exported = []

    for comp in doc.VBProject.VBComponents:
        comp_type = comp.Type
        if comp_type == 100:  # Document module
            continue
        ext = ext_map.get(comp_type)
        if ext is None:
            continue

        out_path = dest / (comp.Name + ext)
        comp.Export(str(out_path.resolve()))
        exported.append(out_path.name)

    return exported


# ---------------------------------------------------------------------------
# Test execution
# ---------------------------------------------------------------------------

def run_test_harness(word, doc, result_file, entry_sub, timeout):
    """
    Run the VBA test harness entry point.
    Sets an environment variable with the result file path so VBA can write to it.
    Returns (success, elapsed_seconds, error_message).
    """
    # Write result path to a known location the VBA harness can read.
    # VBA will use Environ() or read from a temp file.
    result_path = Path(result_file).resolve()
    config_path = result_path.parent / "_harness_config.json"
    with open(config_path, "w", encoding="utf-8") as f:
        json.dump({"result_file": str(result_path)}, f)

    # Also set environment variable for VBA Environ() access
    os.environ["TEST_RESULT_PATH"] = str(result_path)

    # Delete old result file if exists
    if result_path.exists():
        result_path.unlink()

    start = time.time()
    error_msg = None
    success = False

    try:
        # Run the VBA macro
        word.Run(entry_sub)

        elapsed = time.time() - start

        # Check if result file was written
        if result_path.exists():
            success = True
        else:
            error_msg = "Harness did not produce a result file"

    except Exception as e:
        elapsed = time.time() - start
        error_msg = f"COM error running harness: {e}"

    # Enforce timeout
    if elapsed > timeout:
        error_msg = f"Timeout: harness took {elapsed:.1f}s (limit {timeout}s)"
        success = False

    return success, elapsed, error_msg


# ---------------------------------------------------------------------------
# Result parsing and reporting
# ---------------------------------------------------------------------------

def parse_results(result_file):
    """Parse the JSON result file written by the VBA harness."""
    try:
        with open(result_file, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        return {"error": f"Failed to parse results: {e}", "tests": []}


def generate_summary(results, elapsed, iteration, error_msg=None):
    """Generate a concise human-readable summary string."""
    lines = []
    lines.append(f"=== Test Run Summary (iteration {iteration}) ===")
    lines.append(f"Time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    lines.append(f"Elapsed: {elapsed:.2f}s")

    if error_msg:
        lines.append(f"ERROR: {error_msg}")

    tests = results.get("tests", [])
    passed = sum(1 for t in tests if t.get("passed"))
    failed = sum(1 for t in tests if not t.get("passed"))
    total = len(tests)

    lines.append(f"Tests: {total} total, {passed} passed, {failed} failed")

    if failed > 0:
        lines.append("")
        lines.append("Failed tests:")
        for t in tests:
            if not t.get("passed"):
                name = t.get("name", "?")
                msg = t.get("message", "")
                dur = t.get("duration_ms", "?")
                lines.append(f"  FAIL: {name} ({dur}ms) — {msg}")

    # Timeout/crash info
    timed_out = [t for t in tests if t.get("timed_out")]
    if timed_out:
        lines.append("")
        lines.append("Timed-out tests:")
        for t in timed_out:
            lines.append(f"  TIMEOUT: {t.get('name', '?')}")

    lines.append("")
    all_pass = failed == 0 and total > 0 and error_msg is None
    lines.append(f"Result: {'PASS' if all_pass else 'FAIL'}")
    lines.append("=" * 44)

    return "\n".join(lines)


def generate_detailed_report(results, elapsed, iteration, error_msg=None):
    """Generate the full detailed JSON report for machine consumption."""
    tests = results.get("tests", [])
    passed = sum(1 for t in tests if t.get("passed"))
    failed = sum(1 for t in tests if not t.get("passed"))

    report = {
        "iteration": iteration,
        "timestamp": datetime.now().isoformat(),
        "elapsed_seconds": round(elapsed, 3),
        "error": error_msg,
        "summary": {
            "total": len(tests),
            "passed": passed,
            "failed": failed,
            "all_passed": failed == 0 and len(tests) > 0 and error_msg is None
        },
        "tests": tests,
        "raw_harness_output": results
    }
    return report


# ---------------------------------------------------------------------------
# Repair prompt generation
# ---------------------------------------------------------------------------

def generate_repair_prompt(report, config, iteration):
    """Generate a repair prompt file that can be fed to Claude Code CLI."""
    template_path = PROJECT_ROOT / "controller" / "templates" / "repair_prompt.md"
    if template_path.exists():
        with open(template_path, "r", encoding="utf-8") as f:
            template = f.read()
    else:
        template = DEFAULT_REPAIR_TEMPLATE

    # Gather failed tests
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
        full_report_json=json.dumps(report, indent=2)
    )

    return prompt


DEFAULT_REPAIR_TEMPLATE = """# VBA Repair Prompt — Iteration {iteration}/{max_iterations}

## Status
- **{failed_count}** of **{total_count}** tests failed
- Elapsed: {elapsed}s
- Error: {error_message}

## Rules
1. You may ONLY modify files under: {src_dir}/
2. Allowed mutable paths:
{mutable_paths}
3. Do NOT modify any test files, fixtures, expected outputs, harness code, or controller code.
4. Do NOT add new files outside the mutable paths.
5. Focus on fixing the failing tests below.

## Failed Tests
```json
{failed_tests_json}
```

## Full Test Report
```json
{full_report_json}
```

## Instructions
- Read the failing test details above.
- Identify the root cause in the VBA source files under `{src_dir}/`.
- Make minimal, targeted fixes.
- Do not refactor or change code unrelated to the failures.
- After making changes, save all modified files.
"""


# ---------------------------------------------------------------------------
# Main controller loop
# ---------------------------------------------------------------------------

def run_iteration(config, iteration, results_dir):
    """Run one import → test → report cycle. Returns (all_passed, report)."""
    word = None
    doc = None
    result_file = results_dir / f"results_{iteration:03d}.json"

    try:
        visible = config["controller"].get("word_visible", False)
        timeout = config["controller"].get("timeout_seconds", 60)
        entry_sub = config["harness"].get("entry_sub", "RunAllTests")

        word_doc_path = (PROJECT_ROOT / config["word_doc"]).resolve()
        src_dir = (PROJECT_ROOT / config["src_dir"]).resolve()
        harness_dir = (PROJECT_ROOT / "harness").resolve()

        if not word_doc_path.exists():
            return False, {
                "error": f"Word document not found: {word_doc_path}",
                "summary": {"total": 0, "passed": 0, "failed": 0, "all_passed": False},
                "tests": [],
                "iteration": iteration,
                "timestamp": datetime.now().isoformat(),
                "elapsed_seconds": 0
            }

        print(f"  Starting Word (visible={visible})...")
        word = get_word_app(visible=visible)

        print(f"  Opening {word_doc_path.name}...")
        doc = word.Documents.Open(str(word_doc_path))

        print(f"  Importing VBA source from {src_dir}...")
        imported_src = import_vba_modules(word, doc, src_dir)
        print(f"    Imported: {', '.join(imported_src) if imported_src else '(none)'}")

        print(f"  Importing test harness from {harness_dir}...")
        imported_harness = import_harness_modules(word, doc, harness_dir)
        print(f"    Imported: {', '.join(imported_harness) if imported_harness else '(none)'}")

        print(f"  Running test harness ({entry_sub})...")
        success, elapsed, error_msg = run_test_harness(
            word, doc, str(result_file), entry_sub, timeout
        )

        # Parse results
        if result_file.exists():
            results = parse_results(str(result_file))
        else:
            results = {"tests": []}

        report = generate_detailed_report(results, elapsed, iteration, error_msg)
        all_passed = report["summary"]["all_passed"]

        return all_passed, report

    except Exception as e:
        return False, {
            "error": str(e),
            "summary": {"total": 0, "passed": 0, "failed": 0, "all_passed": False},
            "tests": [],
            "iteration": iteration,
            "timestamp": datetime.now().isoformat(),
            "elapsed_seconds": 0
        }

    finally:
        # Clean up Word
        try:
            if doc:
                doc.Close(0)  # wdDoNotSaveChanges
        except Exception:
            pass
        try:
            if word:
                word.Quit()
        except Exception:
            pass


def main():
    config = load_config()
    max_iters = config["controller"]["max_iterations"]
    results_dir = PROJECT_ROOT / config.get("results_dir", "results")
    results_dir.mkdir(parents=True, exist_ok=True)

    locked_paths = load_locked_paths()

    print("=" * 60)
    print("Word VBA Patch Automator — Controller (MVP)")
    print(f"Project: {config['project_name']}")
    print(f"Max iterations: {max_iters}")
    print(f"Timeout: {config['controller']['timeout_seconds']}s")
    print("=" * 60)

    run_reports = []

    for iteration in range(1, max_iters + 1):
        print(f"\n--- Iteration {iteration}/{max_iters} ---")

        # Record commit before this iteration for lock checking
        pre_commit = get_current_commit()

        all_passed, report = run_iteration(config, iteration, results_dir)
        run_reports.append(report)

        # Write detailed JSON report
        report_file = results_dir / f"report_{iteration:03d}.json"
        with open(report_file, "w", encoding="utf-8") as f:
            json.dump(report, f, indent=2)

        # Write human summary
        summary = generate_summary(
            report.get("raw_harness_output", report),
            report["elapsed_seconds"],
            iteration,
            report.get("error")
        )
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
            print(f"\nRepair prompt written to: {prompt_file}")
            print(">>> Run Claude Code CLI against this prompt, then re-run the controller. <<<")

            # In semi-automatic mode, stop here and wait for human
            print("\n[Semi-automatic mode] Waiting for you to run Claude and re-invoke.")
            print(f"  claude -p {prompt_file}")

            # After Claude runs (next iteration), check locked paths
            # For MVP semi-auto, we check on next invocation start
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

    print(f"\nFull run log: {combined_file}")


def check_locks_before_run():
    """Check if any locked files were modified since last run. Call at start."""
    config = load_config()
    locked_paths = load_locked_paths()

    # Read last known commit from a marker file
    marker = PROJECT_ROOT / "results" / ".last_commit"
    if not marker.exists():
        # First run — record current commit
        commit = get_current_commit()
        if commit:
            marker.parent.mkdir(parents=True, exist_ok=True)
            with open(marker, "w") as f:
                f.write(commit)
        return True

    with open(marker, "r") as f:
        last_commit = f.read().strip()

    changed = get_changed_files_since(last_commit)
    violations = check_locked_paths(changed, locked_paths)

    if violations:
        print("LOCKED PATH VIOLATION — the following locked files were modified:")
        for v in violations:
            print(f"  - {v}")
        print("\nReverting is recommended. Aborting.")
        return False

    # Update marker
    commit = get_current_commit()
    if commit:
        with open(marker, "w") as f:
            f.write(commit)
    return True


if __name__ == "__main__":
    if not check_locks_before_run():
        sys.exit(1)
    main()
