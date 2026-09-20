"""
build_portable_zip.py
=====================
Run this script ONCE on your Mac to produce two distributable archives:

  CCWR-Windows-Portable.zip   (Windows, no Python install needed)
  CCWR-Mac-Portable.zip       (Mac, no Python install needed if Python 3 present)

Both zips can be shared via an internal drive. The recipient just unzips and
double-clicks the launcher.  No GitHub access required.

Usage:
  cd "CCWR This-Last Analysis"
  python3 build_portable_zip.py
"""

from __future__ import annotations

import io
import shutil
import subprocess
import sys
import tempfile
import urllib.request
import zipfile
from pathlib import Path

# --- Configuration ----------------------------------------------------------

# Latest Python 3.12 embeddable for Windows 64-bit
# Update the version string here when a newer 3.12 patch is released.
PY_VERSION  = "3.12.9"
PY_EMBED_URL = (
    f"https://www.python.org/ftp/python/{PY_VERSION}/"
    f"python-{PY_VERSION}-embed-amd64.zip"
)

# Files / folders to INCLUDE from the project when building both zips.
INCLUDE = [
    "README.md",
    "app.py",
    "compare_engine.py",
    "excel_reader.py",
    "report_writer.py",
    "ppt_writer.py",
    "price_impact.py",
    "requirements.txt",
    "Run App.bat",
    "Run App.command",
    ".streamlit/config.toml",
    "assets/cisco_template.pptx",
    # Enhanced version (Service Tier Change Impact + New Price Quote Generation)
    "tier_impact_app.py",
    "tier_impact_engine.py",
    "service_tier_classifier.py",
    "Run Tier Impact App.bat",
    "Run Tier Impact App.command",
    "egrid.xlsx",
    # glasia.xlsx (~129 MB) is deliberately NOT bundled -- each user uploads
    # their own copy via the app when Cisco issues an updated price list.
]

# Files / folders to EXCLUDE when scanning the project root.
EXCLUDE_NAMES = {
    "__pycache__", ".venv", ".git", ".github",
    "build_portable_zip.py",      # this script itself
    "CCWR-Windows-Portable.zip",
    "CCWR-Mac-Portable.zip",
    "renewal_comparison.xlsx",    # output files
    "Last Year Sniff.xlsx",
    "This year Sniff.xlsx",
    "Last year quote.xlsx",
    "This year quote.xlsx",
}

PROJECT_DIR = Path(__file__).parent.resolve()
APP_VERSION = "v4"
OUT_WIN     = PROJECT_DIR / f"CCWR-Windows-Portable {APP_VERSION}.zip"
OUT_MAC     = PROJECT_DIR / f"CCWR-Mac-Portable {APP_VERSION}.zip"

# ---------------------------------------------------------------------------

def _progress(msg: str) -> None:
    print(f"  {msg}", flush=True)


def _download_bytes(url: str, label: str) -> bytes:
    _progress(f"Downloading {label} ...")
    with urllib.request.urlopen(url, timeout=120) as r:
        total = int(r.headers.get("Content-Length", 0))
        buf = bytearray()
        chunk = 65536
        downloaded = 0
        while True:
            block = r.read(chunk)
            if not block:
                break
            buf.extend(block)
            downloaded += len(block)
            if total:
                pct = downloaded * 100 // total
                print(f"\r    {pct:3d}%  ({downloaded//1024} KB / {total//1024} KB)", end="", flush=True)
        print()
    return bytes(buf)


def build_windows_zip() -> None:
    print("\n=== Building CCWR-Windows-Portable.zip ===")

    with tempfile.TemporaryDirectory() as tmp_str:
        tmp = Path(tmp_str)
        root = tmp / "CCWR This-Last Analysis"
        root.mkdir()

        # 1. Copy project files (including .streamlit subfolder)
        _progress("Copying project files ...")
        for name in INCLUDE:
            src = PROJECT_DIR / name
            if src.exists():
                dest = root / name
                dest.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(src, dest)

        # 2. Download Python embeddable
        embed_bytes = _download_bytes(PY_EMBED_URL, f"Python {PY_VERSION} embeddable (Windows)")
        py_dir = root / "_python"
        py_dir.mkdir()
        _progress("Extracting Python embeddable ...")
        with zipfile.ZipFile(io.BytesIO(embed_bytes)) as zf:
            zf.extractall(py_dir)

        # 3. Enable site-packages in the embeddable Python by editing the .pth file
        pth_files = list(py_dir.glob("python3*._pth"))
        if not pth_files:
            raise RuntimeError("Could not find ._pth file in embeddable Python!")
        pth_file = pth_files[0]
        content = pth_file.read_text(encoding="utf-8")
        # Uncomment 'import site' AND add Lib/site-packages so installed packages are found
        content = content.replace("#import site", "import site")
        if "Lib/site-packages" not in content:
            content = content.rstrip() + "\nLib/site-packages\n"
        pth_file.write_text(content, encoding="utf-8")

        # 4. Create site-packages directory
        site_pkgs = py_dir / "Lib" / "site-packages"
        site_pkgs.mkdir(parents=True, exist_ok=True)

        # 5. Install Windows-targeted packages using the HOST (Mac) Python + pip.
        #    --platform / --abi flags tell pip to download Windows AMD64 wheels
        #    even though we are running on Mac. --only-binary=:all: ensures we
        #    never try to compile C extensions that would target Mac instead.
        _progress("Installing packages for Windows (cross-platform pip download) ...")
        req_file = root / "requirements.txt"
        result = subprocess.run(
            [
                sys.executable, "-m", "pip", "install",
                "--target", str(site_pkgs),
                "--platform", "win_amd64",
                "--python-version", "3.12",
                "--implementation", "cp",
                "--abi", "cp312",
                "--only-binary=:all:",
                "--upgrade",
                "--quiet",
                "-r", str(req_file),
            ],
            capture_output=True, text=True,
        )
        if result.returncode != 0:
            _progress("Binary-only install failed for some packages; retrying without platform filter for pure-Python packages ...")
            # Some packages are pure Python (no binary wheels), retry without platform filter
            result2 = subprocess.run(
                [
                    sys.executable, "-m", "pip", "install",
                    "--target", str(site_pkgs),
                    "--upgrade",
                    "--quiet",
                    "-r", str(req_file),
                ],
                capture_output=True, text=True,
            )
            if result2.returncode != 0:
                print(result2.stderr[-800:])
                raise RuntimeError("pip install failed — see above.")
        _progress("Packages installed.")

        # 6. Write a robust portable launcher
        # Note: written as bytes with explicit CRLF so cmd.exe parses correctly
        # even though this build script runs on Mac (which defaults to LF).
        portable_bat = (
            "@echo off\r\n"
            "setlocal EnableDelayedExpansion\r\n"
            "cd /d \"%~dp0\"\r\n"
            "\r\n"
            "echo ==================================================\r\n"
            "echo   CCWR Renewal Quote Comparator\r\n"
            "echo ==================================================\r\n"
            "echo.\r\n"
            "\r\n"
            ":: Remove the 'downloaded from internet' block so Windows lets python.exe run\r\n"
            "powershell -ExecutionPolicy Bypass -Command \"Get-ChildItem -Path '%~dp0' -Recurse | Unblock-File\" 2>nul\r\n"
            "\r\n"
            "if not exist \"%~dp0_python\\python.exe\" (\r\n"
            "  echo ERROR: The _python folder is missing or incomplete.\r\n"
            "  echo Please FULLY EXTRACT the zip before running:\r\n"
            "  echo   Right-click the .zip ^> Extract All ^> choose a folder ^> Extract\r\n"
            "  echo.\r\n"
            "  goto :end\r\n"
            ")\r\n"
            "\r\n"
            "echo Starting app...\r\n"
            "echo Your browser will open at http://localhost:8501\r\n"
            "echo Keep this window open while using the app.\r\n"
            "echo Press Ctrl+C in this window to stop.\r\n"
            "echo.\r\n"
            "\r\n"
            "\"%~dp0_python\\python.exe\" -m streamlit run \"%~dp0app.py\"\r\n"
            "\r\n"
            "if errorlevel 1 (\r\n"
            "  echo.\r\n"
            "  echo ====================================================\r\n"
            "  echo  The app stopped or failed to start.\r\n"
            "  echo  Check for error messages above.\r\n"
            "  echo ====================================================\r\n"
            "  echo.\r\n"
            ")\r\n"
            "\r\n"
            ":end\r\n"
            "echo Press any key to close this window...\r\n"
            "pause >nul\r\n"
        )
        (root / "Run App.bat").write_bytes(portable_bat.encode("utf-8"))

        # 6b. Same launcher pattern for the enhanced (tier impact) app -- the
        # plain "Run Tier Impact App.bat" copied via INCLUDE assumes a .venv
        # (GitHub-source use), which doesn't exist in this embedded-Python
        # portable layout, so it's overwritten here too.
        portable_bat_tier = portable_bat.replace(
            "CCWR Renewal Quote Comparator", "CCWR Service Tier Change Impact"
        ).replace(
            "-m streamlit run \"%~dp0app.py\"",
            "-m streamlit run \"%~dp0tier_impact_app.py\"",
        )
        (root / "Run Tier Impact App.bat").write_bytes(portable_bat_tier.encode("utf-8"))

        # 7. Zip everything up
        _progress(f"Creating {OUT_WIN.name} ...")
        with zipfile.ZipFile(OUT_WIN, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as zf:
            for f in sorted(root.rglob("*")):
                if f.is_file():
                    zf.write(f, f.relative_to(tmp))
        size_mb = OUT_WIN.stat().st_size / 1_048_576
        print(f"  Done → {OUT_WIN}  ({size_mb:.0f} MB)")


def build_mac_zip() -> None:
    print("\n=== Building CCWR-Mac-Portable.zip ===")

    with tempfile.TemporaryDirectory() as tmp_str:
        tmp = Path(tmp_str)
        root = tmp / "CCWR This-Last Analysis"
        root.mkdir()

        # Copy project files (including subfolders like .streamlit)
        _progress("Copying project files ...")
        for name in INCLUDE:
            src = PROJECT_DIR / name
            if src.exists():
                dest = root / name
                dest.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(src, dest)

        # Make the launchers executable
        for cmd_name in ("Run App.command", "Run Tier Impact App.command"):
            cmd = root / cmd_name
            if cmd.exists():
                cmd.chmod(0o755)

        # Zip
        _progress(f"Creating {OUT_MAC.name} ...")
        with zipfile.ZipFile(OUT_MAC, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as zf:
            for f in sorted(root.rglob("*")):
                if f.is_file():
                    rel = f.relative_to(tmp)
                    info = zipfile.ZipInfo(str(rel))
                    info.compress_type = zipfile.ZIP_DEFLATED
                    # Preserve executable bit for .command
                    if f.suffix == ".command":
                        info.external_attr = 0o755 << 16
                    else:
                        info.external_attr = 0o644 << 16
                    zf.writestr(info, f.read_bytes())

        size_mb = OUT_MAC.stat().st_size / 1_048_576
        print(f"  Done → {OUT_MAC}  ({size_mb:.1f} MB)")


def main() -> None:
    print("CCWR Portable Distribution Builder")
    print("====================================")
    print(f"Project: {PROJECT_DIR}")

    # Check we can reach python.org
    try:
        urllib.request.urlopen("https://www.python.org", timeout=5)
    except Exception:
        print("\nERROR: Cannot reach python.org. Check your internet connection.")
        sys.exit(1)

    build_windows_zip()
    build_mac_zip()

    print("\n=== All done ===")
    print(f"Share these two files via your internal drive:")
    print(f"  {OUT_WIN}")
    print(f"  {OUT_MAC}")
    print()
    print("Windows: unzip → double-click 'Run App.bat'  (no Python install needed)")
    print("Mac:     unzip → right-click 'Run App.command' → Open")
    print()
    print("Tip: The Windows zip is large (~200 MB) because Python and all packages")
    print("     are bundled. If your internal drive has a size limit, share the")
    print("     regular zip instead (people need Python once, then it just works).")


if __name__ == "__main__":
    main()
