"""pptx2md-gui 便携 EXE 打包脚本。

用法:
    python build_exe.py            # 标准打包（--onedir 模式）
    python build_exe.py --onefile  # 单文件打包（启动较慢，更易被杀软拦截）
    python build_exe.py --clean    # 仅清理构建产物

前置条件:
    pip install -e ".[build]"
"""

import argparse
import hashlib
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
SPEC_BUILD_DIR = BUILD_DIR / SPEC_FILE.stem
ONEFILE_ENV_VAR = "PPTX2MD_GUI_ONEFILE"
REQUIRED_RUNTIME_FILES = (
    "_tkinter.pyd",
    "_ctypes.pyd",
    "_bz2.pyd",
    "_decimal.pyd",
    "tcl86t.dll",
    "tk86t.dll",
    "ffi.dll",
    "libbz2.dll",
    "libmpdec-4.dll",
)
REQUIRED_TCL_DATA_FILES = (
    Path("_tcl_data") / "init.tcl",
    Path("_tk_data") / "tk.tcl",
)


def _iter_unique_prefixes():
    """返回当前解释器相关的唯一前缀目录（已解析绝对路径）。"""
    seen = set()
    for raw_prefix in (sys.prefix, sys.exec_prefix, sys.base_prefix):
        if not raw_prefix:
            continue
        try:
            resolved = Path(raw_prefix).resolve()
        except OSError:
            resolved = Path(os.path.abspath(raw_prefix))
        normalized = os.path.normcase(str(resolved))
        if normalized in seen:
            continue
        seen.add(normalized)
        yield resolved


def _prepare_build_env(use_onefile: bool) -> dict[str, str]:
    """为 PyInstaller 构建准备环境变量，优先使用当前解释器环境中的 DLL。"""
    env = os.environ.copy()
    env[ONEFILE_ENV_VAR] = "1" if use_onefile else "0"

    if sys.platform == "win32":
        preferred_paths = []
        seen = set()
        for prefix in _iter_unique_prefixes():
            for candidate in (
                prefix,
                prefix / "Scripts",
                prefix / "Library" / "bin",
                prefix / "DLLs",
            ):
                if not candidate.is_dir():
                    continue
                resolved = str(candidate)
                normalized = os.path.normcase(resolved)
                if normalized in seen:
                    continue
                seen.add(normalized)
                preferred_paths.append(resolved)

        current_path = env.get("PATH", "")
        env["PATH"] = os.pathsep.join(preferred_paths + ([current_path] if current_path else []))

    return env


def _sha256(path: Path) -> str:
    """计算文件 SHA256，用于校验打包产物是否来自当前构建环境。"""
    digest = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _find_env_runtime_file(name: str) -> Path | None:
    """在当前解释器环境中查找指定运行时文件。"""
    for prefix in _iter_unique_prefixes():
        candidate = prefix / "Library" / "bin" / name
        if candidate.is_file():
            return candidate
    return None


def _text_mentions_path(text: str, path: Path) -> bool:
    """兼容正反斜杠，判断文本中是否包含给定路径。"""
    candidates = {
        str(path).lower(),
        str(path).replace("\\", "/").lower(),
    }
    return any(candidate in text for candidate in candidates)


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


def validate_gui_runtime(output_path: Path, use_onefile: bool) -> bool:
    """校验打包产物中是否包含 GUI 启动所需的关键扩展和 DLL。"""
    if use_onefile:
        analysis_toc = SPEC_BUILD_DIR / "Analysis-00.toc"
        if not analysis_toc.exists():
            print(f"未找到分析文件，无法校验 Tk 运行时: {analysis_toc}")
            return False

        toc_text = analysis_toc.read_text(encoding="utf-8", errors="ignore").lower()
        missing = [name for name in REQUIRED_RUNTIME_FILES if name.lower() not in toc_text]

        missing_data_files = [
            str(rel_path).replace("\\", "/")
            for rel_path in REQUIRED_TCL_DATA_FILES
            if not _text_mentions_path(toc_text, rel_path)
        ]
        if missing_data_files:
            print("GUI 运行时校验失败，缺少以下 Tcl/Tk 脚本文件:")
            for rel_path in missing_data_files:
                print(f"  - {rel_path}")
            return False

        for name in ("tcl86t.dll", "tk86t.dll"):
            expected_source = _find_env_runtime_file(name)
            if expected_source and not _text_mentions_path(toc_text, expected_source):
                print(f"GUI 运行时校验失败，{name} 不是从当前构建环境收集的: {expected_source}")
                return False
    else:
        internal_dir = output_path / "_internal"
        if not internal_dir.exists():
            print(f"未找到 onedir 运行时目录: {internal_dir}")
            return False

        packaged_names = {path.name.lower() for path in internal_dir.iterdir()}
        missing = [name for name in REQUIRED_RUNTIME_FILES if name.lower() not in packaged_names]

        missing_data_files = [
            str(rel_path).replace("\\", "/")
            for rel_path in REQUIRED_TCL_DATA_FILES
            if not (internal_dir / rel_path).exists()
        ]
        if missing_data_files:
            print("GUI 运行时校验失败，缺少以下 Tcl/Tk 脚本文件:")
            for rel_path in missing_data_files:
                print(f"  - {rel_path}")
            return False

        for name in ("tcl86t.dll", "tk86t.dll"):
            expected_source = _find_env_runtime_file(name)
            packaged_file = internal_dir / name
            if not expected_source or not packaged_file.is_file():
                continue
            if _sha256(expected_source) != _sha256(packaged_file):
                print(f"GUI 运行时校验失败，{name} 不是当前构建环境中的版本:")
                print(f"  - 期望来源: {expected_source}")
                print(f"  - 实际产物: {packaged_file}")
                return False

    if missing:
        print("GUI 运行时校验失败，缺少以下关键文件:")
        for name in missing:
            print(f"  - {name}")
        return False

    print("GUI 运行时校验通过。")
    return True


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

    env = _prepare_build_env(use_onefile)

    result = subprocess.run(cmd, cwd=str(PROJECT_ROOT), env=env)
    if result.returncode == 0:
        output_path = DIST_DIR / f"{app_name}.exe" if use_onefile else DIST_DIR / app_name
        if not validate_gui_runtime(output_path, use_onefile=use_onefile):
            print("\n构建产物缺少 GUI 运行时文件，请检查 spec 配置。")
            return False
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
