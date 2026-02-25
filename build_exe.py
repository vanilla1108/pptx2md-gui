"""pptx2md-gui 便携 EXE 打包脚本。

用法:
    python build_exe.py            # 标准打包（--onedir 模式）
    python build_exe.py --onefile  # 单文件打包（启动较慢，更易被杀软拦截）
    python build_exe.py --clean    # 仅清理构建产物

前置条件:
    pip install -e ".[build]"
"""

import argparse
import os
import shutil
import subprocess
import sys
import tomllib
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent
BUILD_DIR = PROJECT_ROOT / "build"
DIST_DIR = PROJECT_ROOT / "dist"
SPEC_FILE = PROJECT_ROOT / "pptx2md_gui.spec"
ONEFILE_ENV_VAR = "PPTX2MD_GUI_ONEFILE"


def get_app_version() -> str:
    """从 pyproject.toml 读取版本号。"""
    with open(PROJECT_ROOT / "pyproject.toml", "rb") as f:
        pyproject = tomllib.load(f)
    return pyproject["tool"]["poetry"]["version"]


def clean_build_artifacts():
    """清理残留的 build 和 dist 目录。"""
    for d in [BUILD_DIR, DIST_DIR]:
        if d.exists():
            print(f"正在清理 {d} ...")
            shutil.rmtree(d, ignore_errors=True)

    # 清理源码目录中的 __pycache__
    for pkg in ["pptx2md", "pptx2md_gui"]:
        pkg_dir = PROJECT_ROOT / pkg
        for cache_dir in pkg_dir.rglob("__pycache__"):
            shutil.rmtree(cache_dir, ignore_errors=True)

    print("构建产物已清理。")


def build(use_onefile: bool = False):
    """使用 spec 文件执行 PyInstaller 打包。"""
    version = get_app_version()
    app_name = f"pptx2md-gui-{version}"

    cmd = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--clean",
        "--noconfirm",
        str(SPEC_FILE),
    ]

    if use_onefile:
        print("警告: --onefile 模式更容易被杀软拦截。")
        print("如果构建失败，请先尝试不带 --onefile 的标准模式。\n")

    print(f"版本: {version}")
    print(f"执行命令: {' '.join(cmd)}\n")

    env = os.environ.copy()
    env[ONEFILE_ENV_VAR] = "1" if use_onefile else "0"

    result = subprocess.run(cmd, cwd=str(PROJECT_ROOT), env=env)
    if result.returncode == 0:
        output_path = DIST_DIR / f"{app_name}.exe" if use_onefile else DIST_DIR / app_name
        print(f"\n构建成功！输出: {output_path}")
        return True

    print(f"\n构建失败（退出码 {result.returncode}）。")
    return False


def main():
    parser = argparse.ArgumentParser(description="构建 pptx2md-gui 便携 EXE")
    parser.add_argument(
        "--onefile",
        action="store_true",
        help="打包为单文件 EXE（启动较慢，更易被杀软拦截）",
    )
    parser.add_argument(
        "--clean",
        action="store_true",
        help="仅清理构建产物，不执行打包",
    )
    args = parser.parse_args()

    if args.clean:
        clean_build_artifacts()
        return

    # 前置检查
    try:
        import PyInstaller  # noqa: F401
    except ImportError:
        print("未安装 PyInstaller，请执行:")
        print("  pip install pyinstaller")
        sys.exit(1)

    if sys.platform == "win32":
        try:
            import win32com.client  # noqa: F401
        except ImportError:
            print("Windows 下需要 pywin32 以支持 .ppt 转换和 WMF COM 回退。")
            print("请执行:")
            print("  pip install pywin32")
            sys.exit(1)

    if not SPEC_FILE.exists():
        print(f"未找到 spec 文件: {SPEC_FILE}")
        sys.exit(1)

    clean_build_artifacts()

    success = build(use_onefile=args.onefile)
    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
