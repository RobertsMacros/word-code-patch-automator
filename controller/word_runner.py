"""
Word Runner — subprocess worker for COM automation.

Called by the controller as a child process. This is the process boundary
that allows the controller to enforce a real timeout: if this process
hangs inside a Word COM call, the controller can kill it and Word.

Usage (called by controller, not directly):
    python word_runner.py <config_json_path>

The config_json_path points to a temporary JSON file with run parameters.
Results are written to the path specified in that config.

Exit codes:
    0 = completed (check result JSON for pass/fail)
    1 = error (error details in result JSON)
"""

import json
import os
import sys
import time
from datetime import datetime
from pathlib import Path


def get_word_app(visible=False):
    """Launch a fresh Word instance via COM."""
    import win32com.client
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = visible
    word.DisplayAlerts = 0  # wdAlertsNone
    return word


def import_vba_modules(doc, src_dir):
    """Import .bas/.cls/.frm from src_dir into doc, replacing existing."""
    vb_project = doc.VBProject
    src_path = Path(src_dir)
    valid_ext = {".bas", ".cls", ".frm"}
    imported = []

    for fpath in sorted(src_path.iterdir()):
        if fpath.is_dir() or fpath.suffix.lower() not in valid_ext:
            continue

        mod_name = fpath.stem

        # Remove existing module if present (skip document modules type=100)
        try:
            existing = vb_project.VBComponents(mod_name)
            if existing.Type != 100:
                vb_project.VBComponents.Remove(existing)
                time.sleep(0.3)
        except Exception:
            pass

        vb_project.VBComponents.Import(str(fpath.resolve()))
        imported.append(fpath.name)

    return imported


def run(run_config):
    """
    Execute one test run:
      1. Open the host document
      2. Import VBA source + harness modules
      3. Run the harness entry sub (which iterates fixtures internally)
      4. Read the result file written by the harness
      5. Write a combined result JSON for the controller

    The VBA harness receives fixture paths and result path via a config
    file at a well-known location.
    """
    word = None
    host_doc = None

    result = {
        "status": "error",
        "error": None,
        "elapsed_seconds": 0,
        "fixtures": [],
        "tests": [],
        "timestamp": datetime.now().isoformat()
    }

    start = time.time()

    try:
        visible = run_config.get("word_visible", False)
        host_doc_path = run_config["host_doc_path"]
        src_dir = run_config["src_dir"]
        harness_dir = run_config["harness_dir"]
        entry_sub = run_config.get("entry_sub", "RunAllTests")
        result_file = run_config["result_file"]
        fixtures = run_config.get("fixtures", [])

        # Write harness config to TEMP so VBA can find it reliably
        temp_dir = os.environ.get("TEMP", os.environ.get("TMP", "."))
        harness_config_path = os.path.join(temp_dir, "_harness_config.json")
        harness_config = {
            "result_file": result_file,
            "fixtures": fixtures
        }
        with open(harness_config_path, "w", encoding="utf-8") as f:
            json.dump(harness_config, f)

        # Also set env vars for VBA Environ() access
        os.environ["TEST_RESULT_PATH"] = result_file
        os.environ["TEST_HARNESS_CONFIG"] = harness_config_path

        # Delete old result file
        if os.path.exists(result_file):
            os.unlink(result_file)

        # Start Word
        word = get_word_app(visible=visible)

        # Open host document
        if not os.path.exists(host_doc_path):
            result["error"] = f"Host document not found: {host_doc_path}"
            return result

        host_doc = word.Documents.Open(host_doc_path)

        # Import VBA source modules
        imported_src = import_vba_modules(host_doc, src_dir)

        # Import harness modules
        imported_harness = import_vba_modules(host_doc, harness_dir)

        # Run the test harness
        word.Run(entry_sub)

        elapsed = time.time() - start
        result["elapsed_seconds"] = round(elapsed, 3)

        # Read the result file written by the harness
        if os.path.exists(result_file):
            with open(result_file, "r", encoding="utf-8") as f:
                harness_output = json.load(f)
            result["status"] = "completed"
            result["tests"] = harness_output.get("tests", [])
            result["fixtures"] = harness_output.get("fixtures", [])
            result["harness_output"] = harness_output
        else:
            result["error"] = "Harness did not produce a result file"

    except Exception as e:
        result["elapsed_seconds"] = round(time.time() - start, 3)
        result["error"] = str(e)

    finally:
        try:
            if host_doc:
                host_doc.Close(0)  # wdDoNotSaveChanges
        except Exception:
            pass
        try:
            if word:
                word.Quit()
        except Exception:
            pass

    return result


def main():
    if len(sys.argv) < 2:
        print("Usage: word_runner.py <run_config.json>", file=sys.stderr)
        sys.exit(1)

    config_path = sys.argv[1]
    with open(config_path, "r", encoding="utf-8") as f:
        run_config = json.load(f)

    output_path = run_config["runner_output_file"]

    result = run(run_config)

    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(result, f, indent=2)

    sys.exit(0 if result["status"] == "completed" else 1)


if __name__ == "__main__":
    main()
