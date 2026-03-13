"""
VBA Import / Export utility — standalone script for manual use.

Usage:
    python vba_io.py import          Import src/ modules into the Word document
    python vba_io.py export          Export VBA modules from the Word document to src/
    python vba_io.py export-all      Export all modules including harness/test modules
    python vba_io.py roundtrip       Export then import (useful for normalising)

Runs inside Windows with pywin32 installed.
"""

import json
import sys
import time
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = SCRIPT_DIR.parent


def load_config():
    with open(PROJECT_ROOT / "project.json", "r", encoding="utf-8") as f:
        return json.load(f)


def get_word_app(visible=False):
    import win32com.client
    try:
        word = win32com.client.GetActiveObject("Word.Application")
    except Exception:
        word = win32com.client.Dispatch("Word.Application")
    word.Visible = visible
    word.DisplayAlerts = 0
    return word


def import_modules(word, doc, src_dir):
    """Import .bas/.cls/.frm files from src_dir, replacing existing modules."""
    vb_project = doc.VBProject
    src_path = Path(src_dir)
    ext_types = {".bas": 1, ".cls": 2, ".frm": 3}
    imported = []

    for fpath in sorted(src_path.iterdir()):
        ext = fpath.suffix.lower()
        if ext not in ext_types:
            continue

        mod_name = fpath.stem

        # Remove existing (skip document modules type=100)
        try:
            existing = vb_project.VBComponents(mod_name)
            if existing.Type != 100:
                vb_project.VBComponents.Remove(existing)
                time.sleep(0.3)  # Brief pause for COM stability
        except Exception:
            pass

        vb_project.VBComponents.Import(str(fpath.resolve()))
        imported.append(fpath.name)
        print(f"  Imported: {fpath.name}")

    return imported


def export_modules(doc, dest_dir, include_tests=False):
    """Export VBA components to dest_dir."""
    dest = Path(dest_dir)
    dest.mkdir(parents=True, exist_ok=True)
    ext_map = {1: ".bas", 2: ".cls", 3: ".frm"}
    exported = []

    for comp in doc.VBProject.VBComponents:
        comp_type = comp.Type
        if comp_type == 100:  # Document module — skip
            continue

        # Optionally skip test/harness modules
        if not include_tests:
            if comp.Name.startswith("Test_") or comp.Name == "TestHarness":
                continue

        ext = ext_map.get(comp_type)
        if ext is None:
            continue

        out_path = dest / (comp.Name + ext)
        comp.Export(str(out_path.resolve()))
        exported.append(out_path.name)
        print(f"  Exported: {out_path.name}")

    return exported


def main():
    if len(sys.argv) < 2:
        print("Usage: python vba_io.py [import|export|export-all|roundtrip]")
        sys.exit(1)

    action = sys.argv[1].lower()
    config = load_config()
    visible = "--visible" in sys.argv

    word_doc_path = (PROJECT_ROOT / config["word_doc"]).resolve()
    src_dir = (PROJECT_ROOT / config["src_dir"]).resolve()

    if not word_doc_path.exists():
        print(f"Error: Document not found: {word_doc_path}")
        sys.exit(1)

    print(f"Opening Word (visible={visible})...")
    word = get_word_app(visible=visible)

    try:
        doc = word.Documents.Open(str(word_doc_path))

        if action == "import":
            print(f"Importing from {src_dir}...")
            imported = import_modules(word, doc, src_dir)
            doc.Save()
            print(f"Done. Imported {len(imported)} module(s).")

        elif action == "export":
            print(f"Exporting to {src_dir}...")
            exported = export_modules(doc, src_dir, include_tests=False)
            print(f"Done. Exported {len(exported)} module(s).")

        elif action == "export-all":
            print(f"Exporting all modules to {src_dir}...")
            exported = export_modules(doc, src_dir, include_tests=True)
            print(f"Done. Exported {len(exported)} module(s).")

        elif action == "roundtrip":
            print("Exporting...")
            export_modules(doc, src_dir, include_tests=False)
            print("Re-importing...")
            import_modules(word, doc, src_dir)
            doc.Save()
            print("Roundtrip complete.")

        else:
            print(f"Unknown action: {action}")
            sys.exit(1)

        doc.Close(0)

    finally:
        try:
            word.Quit()
        except Exception:
            pass


if __name__ == "__main__":
    main()
