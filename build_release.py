from __future__ import annotations

import os
import shutil
import subprocess
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parent
APP_NAME = "平谷周报生成工具"
BUILD_ROOT = ROOT / "work" / "build"
PYINSTALLER_CONFIG = BUILD_ROOT / "pyinstaller_config"
DIST_ROOT = ROOT / "release"
RELEASE_DIR = DIST_ROOT / APP_NAME


def copy_tree(src: Path, dst: Path) -> None:
    if not src.exists():
        return
    if dst.exists():
        shutil.rmtree(dst)
    shutil.copytree(src, dst)


def main() -> int:
    BUILD_ROOT.mkdir(parents=True, exist_ok=True)
    PYINSTALLER_CONFIG.mkdir(parents=True, exist_ok=True)
    DIST_ROOT.mkdir(parents=True, exist_ok=True)

    env = os.environ.copy()
    env.update(
        {
            "TEMP": str(BUILD_ROOT / "tmp"),
            "TMP": str(BUILD_ROOT / "tmp"),
            "TMPDIR": str(BUILD_ROOT / "tmp"),
            "PYINSTALLER_CONFIG_DIR": str(PYINSTALLER_CONFIG),
            "PIP_CACHE_DIR": str(BUILD_ROOT / "pip_cache"),
        }
    )
    Path(env["TEMP"]).mkdir(parents=True, exist_ok=True)
    Path(env["PIP_CACHE_DIR"]).mkdir(parents=True, exist_ok=True)

    cmd = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--clean",
        "--windowed",
        "--name",
        APP_NAME,
        "--distpath",
        str(DIST_ROOT),
        "--workpath",
        str(BUILD_ROOT / "pyinstaller_work"),
        "--specpath",
        str(BUILD_ROOT),
        str(ROOT / "app.py"),
    ]
    subprocess.run(cmd, cwd=ROOT, env=env, check=True)

    copy_tree(ROOT / "docs", RELEASE_DIR / "docs")
    copy_tree(ROOT / "input", RELEASE_DIR / "input")
    (RELEASE_DIR / "output").mkdir(exist_ok=True)
    (RELEASE_DIR / "work").mkdir(exist_ok=True)
    shutil.copy2(ROOT / "README.md", RELEASE_DIR / "README.md")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
