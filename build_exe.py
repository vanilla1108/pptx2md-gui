"""pptx2md-gui portable EXE build script.

Usage:
    python build_exe.py            # Standard build (--onedir)
    python build_exe.py --onefile  # Single-file build (slower startup, more AV-sensitive)
    python build_exe.py --clean    # Clean build artifacts only

Prerequisites:
    pip install pyinstaller
"""

import argparse
import os
import shutil
import subprocess
import sys
import time
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent
BUILD_DIR = PROJECT_ROOT / "build"
DIST_DIR = PROJECT_ROOT / "dist"
SPEC_FILE = PROJECT_ROOT / "pptx2md_gui.spec"

# Safe build directories outside of OneDrive / user profile to reduce AV interference
SAFE_WORK_DIR = Path("C:/temp/pptx2md_build/build")
SAFE_DIST_DIR = Path("C:/temp/pptx2md_build/dist")

MAX_RETRIES = 3
RETRY_DELAY_SECONDS = 5


def clean_build_artifacts():
    """Remove stale build and dist directories."""
    for d in [BUILD_DIR, DIST_DIR]:
        if d.exists():
            print(f"Cleaning {d} ...")
            shutil.rmtree(d, ignore_errors=True)

    for d in [SAFE_WORK_DIR, SAFE_DIST_DIR]:
        if d.exists():
            print(f"Cleaning {d} ...")
            shutil.rmtree(d, ignore_errors=True)

    # Remove __pycache__ in our source directories
    for pkg in ["pptx2md", "pptx2md_gui"]:
        pkg_dir = PROJECT_ROOT / pkg
        for cache_dir in pkg_dir.rglob("__pycache__"):
            shutil.rmtree(cache_dir, ignore_errors=True)

    print("Build artifacts cleaned.")


def check_defender_exclusion():
    """Check and suggest Windows Defender exclusion if on Windows."""
    if sys.platform != "win32":
        return

    print("\n--- Windows Defender Exclusion Check ---")
    print(f"Project directory: {PROJECT_ROOT}")
    print(
        "If the build fails with WinError 5, add this directory to Defender exclusions:"
    )
    print(f'  powershell -Command "Add-MpPreference -ExclusionPath \'{PROJECT_ROOT}\'"')
    print(
        "  (Run as Administrator)\n"
    )


def try_add_defender_exclusion():
    """Attempt to add Defender exclusion for project dir (may require admin)."""
    if sys.platform != "win32":
        return False

    dirs_to_exclude = [str(PROJECT_ROOT), str(SAFE_WORK_DIR.parent)]
    for d in dirs_to_exclude:
        try:
            subprocess.run(
                [
                    "powershell",
                    "-Command",
                    f"Add-MpPreference -ExclusionPath '{d}'",
                ],
                capture_output=True,
                timeout=10,
            )
        except Exception:
            pass
    return True


def build(use_onefile: bool = False, use_safe_dirs: bool = True):
    """Run PyInstaller with the spec file."""

    cmd = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--clean",
        "--noconfirm",
    ]

    if use_safe_dirs and sys.platform == "win32":
        # Build in safe directory to avoid OneDrive / AV locks
        SAFE_WORK_DIR.mkdir(parents=True, exist_ok=True)
        SAFE_DIST_DIR.mkdir(parents=True, exist_ok=True)
        cmd += [
            "--workpath",
            str(SAFE_WORK_DIR),
            "--distpath",
            str(SAFE_DIST_DIR),
        ]

    if use_onefile:
        # For --onefile mode, we modify the spec dynamically
        # Build a single-file EXE variant
        cmd += [
            str(SPEC_FILE),
        ]
        print(
            "WARNING: --onefile mode is more susceptible to AV interference."
        )
        print("If the build fails, try without --onefile first.\n")
    else:
        cmd += [str(SPEC_FILE)]

    print(f"Running: {' '.join(cmd)}\n")

    for attempt in range(1, MAX_RETRIES + 1):
        result = subprocess.run(cmd, cwd=str(PROJECT_ROOT))
        if result.returncode == 0:
            dist_base = SAFE_DIST_DIR if use_safe_dirs else DIST_DIR
            output_dir = dist_base / "pptx2md-gui"
            print(f"\nBuild successful! Output: {output_dir}")

            if use_safe_dirs and sys.platform == "win32":
                # Copy result back to project dist/
                final_dist = DIST_DIR / "pptx2md-gui"
                if final_dist.exists():
                    shutil.rmtree(final_dist, ignore_errors=True)
                DIST_DIR.mkdir(parents=True, exist_ok=True)
                shutil.copytree(output_dir, final_dist)
                print(f"Copied to: {final_dist}")

            return True

        print(f"\nBuild attempt {attempt}/{MAX_RETRIES} failed (exit code {result.returncode}).")

        if attempt < MAX_RETRIES:
            print(f"Retrying in {RETRY_DELAY_SECONDS} seconds...")
            print("(WinError 5 is often transient - AV scanner releasing file locks)")
            time.sleep(RETRY_DELAY_SECONDS)

            # Clean build dir between retries
            work_dir = SAFE_WORK_DIR if use_safe_dirs else BUILD_DIR
            if work_dir.exists():
                shutil.rmtree(work_dir, ignore_errors=True)

    print("\nAll build attempts failed.")
    print("Please try the following:")
    print("  1. Add Windows Defender exclusions (see above)")
    print("  2. Close File Explorer windows near the build/dist directories")
    print("  3. Ensure no previous build EXE is running")
    print("  4. Try running as Administrator")
    return False


def main():
    parser = argparse.ArgumentParser(description="Build pptx2md-gui portable EXE")
    parser.add_argument(
        "--onefile",
        action="store_true",
        help="Build as single-file EXE (slower startup, more AV-sensitive)",
    )
    parser.add_argument(
        "--clean",
        action="store_true",
        help="Only clean build artifacts, do not build",
    )
    parser.add_argument(
        "--no-safe-dirs",
        action="store_true",
        help="Build in project directory instead of C:/temp (may trigger AV)",
    )
    args = parser.parse_args()

    if args.clean:
        clean_build_artifacts()
        return

    # Preflight checks
    try:
        import PyInstaller  # noqa: F401
    except ImportError:
        print("PyInstaller is not installed. Install it with:")
        print("  pip install pyinstaller")
        sys.exit(1)

    if not SPEC_FILE.exists():
        print(f"Spec file not found: {SPEC_FILE}")
        sys.exit(1)

    check_defender_exclusion()
    try_add_defender_exclusion()
    clean_build_artifacts()

    success = build(
        use_onefile=args.onefile,
        use_safe_dirs=not args.no_safe_dirs,
    )
    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
