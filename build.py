#!/usr/bin/env python3
"""Build script for creating standalone executable."""

import subprocess
import shutil
import sys
from pathlib import Path


def clean():
    """Remove build artifacts."""
    print("Cleaning build artifacts...")
    for path in ["build", "dist"]:
        if Path(path).exists():
            shutil.rmtree(path)
            print(f"  Removed {path}/")


def build():
    """Build the executable using PyInstaller."""
    print("Building executable...")
    result = subprocess.run(
        ["pyinstaller", "pptmod.spec", "--log-level=WARN"],
        capture_output=False
    )
    
    if result.returncode == 0:
        exe_path = Path("dist/pptmod.exe")
        if exe_path.exists():
            size_mb = exe_path.stat().st_size / (1024 * 1024)
            print(f"\n{'=' * 40}")
            print("BUILD SUCCESSFUL!")
            print(f"{'=' * 40}")
            print(f"\nExecutable: {exe_path}")
            print(f"Size: {size_mb:.2f} MB")
            print(f"\nRun: .\\dist\\pptmod.exe")
        else:
            print("Error: Executable not found")
            sys.exit(1)
    else:
        print("Build failed!")
        sys.exit(result.returncode)


def main():
    """Main entry point."""
    if len(sys.argv) > 1 and sys.argv[1] == "--clean":
        clean()
    
    build()


if __name__ == "__main__":
    main()
