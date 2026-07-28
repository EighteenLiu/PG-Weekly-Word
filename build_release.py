from __future__ import annotations

import os
import shutil
import stat
import subprocess
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parent
APP_NAME = "平谷周报生成工具"
BUILD_ROOT = ROOT / "work" / "build"
PYINSTALLER_CONFIG = BUILD_ROOT / "pyinstaller_config"
STAGE_DIST = BUILD_ROOT / "dist"
DIST_ROOT = ROOT / "release"
RELEASE_DIR = DIST_ROOT / APP_NAME
STAGE_APP_DIR = STAGE_DIST / APP_NAME


def make_writable(path: Path) -> None:
    try:
        path.chmod(stat.S_IREAD | stat.S_IWRITE)
    except OSError:
        pass


def remove_tree(path: Path) -> None:
    if not path.exists():
        return
    for item in path.rglob("*"):
        make_writable(item)
    make_writable(path)
    shutil.rmtree(path)


def copy_file_unlocked(src: str, dst: str) -> str:
    shutil.copyfile(src, dst)
    make_writable(Path(dst))
    return dst


def copy_tree_unlocked(src: Path, dst: Path) -> None:
    if not src.exists():
        return
    if dst.exists():
        remove_tree(dst)
    shutil.copytree(src, dst, copy_function=copy_file_unlocked)
    for item in dst.rglob("*"):
        make_writable(item)


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
            str(STAGE_DIST),
        "--workpath",
        str(BUILD_ROOT / "pyinstaller_work"),
        "--specpath",
        str(BUILD_ROOT),
        str(ROOT / "app.py"),
    ]
    remove_tree(STAGE_APP_DIR)
    subprocess.run(cmd, cwd=ROOT, env=env, check=True)

    remove_tree(RELEASE_DIR)
    copy_tree_unlocked(STAGE_APP_DIR, RELEASE_DIR)
    copy_tree_unlocked(ROOT / "docs", RELEASE_DIR / "docs")
    copy_tree_unlocked(ROOT / "input", RELEASE_DIR / "input")
    (RELEASE_DIR / "output").mkdir(exist_ok=True)
    (RELEASE_DIR / "work").mkdir(exist_ok=True)
    copy_file_unlocked(str(ROOT / "README.md"), str(RELEASE_DIR / "README.md"))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
