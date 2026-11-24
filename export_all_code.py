"""
Export All Code — dumps every project source file into a single .txt file.

Usage:
    python export_all_code.py
    python export_all_code.py --output my_dump.txt

Output: project_code_export.txt (default) in the project root.
Each file is separated by a clear header with the relative path,
so the entire codebase can be read, pasted, or audited in one go.
"""

import os
import sys
from datetime import datetime
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent

# Files and directories to include, in order.
# Paths are relative to PROJECT_ROOT.
INCLUDE = [
    "project.json",
    "CLAUDE.md",
    ".gitattributes",
    ".gitignore",
    "requirements.txt",
    "controller/controller.py",
    "controller/word_runner.py",
    "controller/vba_io.py",
    "controller/__init__.py",
    "controller/templates/repair_prompt.md",
    "harness/TestHarness.bas",
    "harness/Test_Smoke.bas",
    "src/",          # all files in src/
    "README.md",
]

# Extensions to include when scanning directories
SOURCE_EXTENSIONS = {
    ".py", ".bas", ".cls", ".frm", ".json", ".md", ".txt", ".cfg", ".ini",
    ".yml", ".yaml", ".toml",
}

# Files/dirs to always skip
SKIP = {
    "__pycache__", ".git", "results", "node_modules", ".venv", "venv",
    "export_all_code.py",  # don't include ourselves
}


def collect_files():
    """Build ordered list of files to export."""
    files = []
    seen = set()

    for entry in INCLUDE:
        path = PROJECT_ROOT / entry

        if path.is_file():
            rel = path.relative_to(PROJECT_ROOT)
            if str(rel) not in seen:
                files.append(path)
                seen.add(str(rel))

        elif path.is_dir():
            for child in sorted(path.rglob("*")):
                if child.is_file() and child.suffix.lower() in SOURCE_EXTENSIONS:
                    if not any(skip in child.parts for skip in SKIP):
                        rel = child.relative_to(PROJECT_ROOT)
                        if str(rel) not in seen:
                            files.append(child)
                            seen.add(str(rel))

    return files


def export(output_path):
    files = collect_files()

    separator = "=" * 78

    lines = []
    lines.append(separator)
    lines.append(f"  PROJECT CODE EXPORT — word-code-patch-automator")
    lines.append(f"  Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    lines.append(f"  Files: {len(files)}")
    lines.append(separator)
    lines.append("")

    # Table of contents
    lines.append("TABLE OF CONTENTS")
    lines.append("-" * 40)
    for i, fpath in enumerate(files, 1):
        rel = fpath.relative_to(PROJECT_ROOT)
        lines.append(f"  {i:>2}. {rel}")
    lines.append("")
    lines.append("")

    # Each file
    for i, fpath in enumerate(files, 1):
        rel = fpath.relative_to(PROJECT_ROOT)
        lines.append(separator)
        lines.append(f"  FILE {i}: {rel}")
        lines.append(separator)
        lines.append("")

        try:
            content = fpath.read_text(encoding="utf-8")
            # Strip trailing whitespace per line, preserve structure
            for line in content.splitlines():
                lines.append(line.rstrip())
            # Ensure file content ends with blank line
            if lines[-1] != "":
                lines.append("")
        except Exception as e:
            lines.append(f"[ERROR reading file: {e}]")
            lines.append("")

    lines.append(separator)
    lines.append("  END OF EXPORT")
    lines.append(separator)

    output_path.write_text("\n".join(lines), encoding="utf-8", newline="\n")
    return len(files)


def main():
    output_name = "project_code_export.txt"

    # Allow --output override
    for i, arg in enumerate(sys.argv):
        if arg == "--output" and i + 1 < len(sys.argv):
            output_name = sys.argv[i + 1]

    output_path = PROJECT_ROOT / output_name
    count = export(output_path)
    print(f"Exported {count} files to {output_path}")


if __name__ == "__main__":
    main()
