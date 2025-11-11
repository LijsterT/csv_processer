"""Helper script to bundle the Personalized CSV Creator with PyInstaller."""

import argparse
import importlib.util
import shutil
from pathlib import Path
from typing import Optional

PROJECT_ROOT = Path(__file__).resolve().parent
APP_ENTRY_POINT = PROJECT_ROOT / "app.py"
DEFAULT_NAME = "PersonalizedCSVCreator"


def ensure_pyinstaller() -> None:
    """Ensure PyInstaller is installed before attempting to build."""
    if importlib.util.find_spec("PyInstaller") is None:
        raise SystemExit(
            "PyInstaller is required. Install it with 'pip install pyinstaller' and try again."
        )


def clean_build_artifacts(name: str) -> None:
    """Remove PyInstaller build artifacts for a fresh build."""
    build_dir = PROJECT_ROOT / "build"
    dist_dir = PROJECT_ROOT / "dist"
    spec_file = PROJECT_ROOT / f"{name}.spec"

    if build_dir.exists():
        shutil.rmtree(build_dir)
    if dist_dir.exists():
        shutil.rmtree(dist_dir)
    if spec_file.exists():
        spec_file.unlink()


def build_executable(
    name: str,
    onefile: bool,
    windowed: bool,
    clean: bool,
    icon: Optional[Path],
) -> None:
    """Invoke PyInstaller with the requested options."""
    ensure_pyinstaller()

    if not APP_ENTRY_POINT.exists():
        raise SystemExit(f"Could not locate application entry point: {APP_ENTRY_POINT}")

    if clean:
        clean_build_artifacts(name)

    from PyInstaller.__main__ import run

    args = [f"--name={name}", "--clean"]
    if onefile:
        args.append("--onefile")
    if windowed:
        args.append("--noconsole")
    if icon is not None:
        if not icon.exists():
            raise SystemExit(f"Icon file not found: {icon}")
        args.append(f"--icon={icon}")
    args.append(str(APP_ENTRY_POINT))

    run(args)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Bundle the Personalized CSV Creator into a standalone executable."
    )
    parser.add_argument(
        "--name",
        default=DEFAULT_NAME,
        help=f"Name for the generated executable (default: {DEFAULT_NAME}).",
    )
    parser.add_argument(
        "--onedir",
        action="store_true",
        help="Create a directory-based bundle instead of a single file.",
    )
    parser.add_argument(
        "--console",
        action="store_true",
        help="Keep the console window (omit --noconsole).",
    )
    parser.add_argument(
        "--skip-clean",
        action="store_true",
        help="Skip removing previous PyInstaller build artifacts before packaging.",
    )
    parser.add_argument(
        "--icon",
        type=Path,
        help="Optional path to an icon file to embed in the executable (platform dependent).",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    icon_path = args.icon.expanduser().resolve() if args.icon else None

    build_executable(
        name=args.name,
        onefile=not args.onedir,
        windowed=not args.console,
        clean=not args.skip_clean,
        icon=icon_path,
    )


if __name__ == "__main__":
    main()
