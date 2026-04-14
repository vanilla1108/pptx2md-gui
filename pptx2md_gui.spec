# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec file for pptx2md-gui portable build."""

import os
import re
import sys
import tomllib
from pathlib import Path

from PyInstaller.utils.hooks import collect_data_files, collect_dynamic_libs, collect_submodules
from PyInstaller.utils.hooks.tcl_tk import tcltk_info
from PyInstaller.utils.win32.versioninfo import (
    FixedFileInfo,
    StringFileInfo,
    StringStruct,
    StringTable,
    VarFileInfo,
    VarStruct,
    VSVersionInfo,
)

block_cipher = None
IS_ONEFILE = os.environ.get("PPTX2MD_GUI_ONEFILE", "").strip() == "1"

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------

PROJECT_ROOT = SPECPATH  # SPECPATH is already the directory containing the spec file

# ---------------------------------------------------------------------------
# Version
# ---------------------------------------------------------------------------

with open(os.path.join(PROJECT_ROOT, "pyproject.toml"), "rb") as _f:
    _pyproject = tomllib.load(_f)
APP_VERSION = _pyproject["tool"]["poetry"]["version"]

# Parse "0.1.0b5" -> (0, 1, 0, 5) for Windows version resource
_m = re.match(r"(\d+)\.(\d+)\.(\d+)(?:[ab](\d+))?", APP_VERSION)
_ver_tuple = tuple(int(x or 0) for x in _m.groups())

_version_info = VSVersionInfo(
    ffi=FixedFileInfo(
        filevers=_ver_tuple,
        prodvers=_ver_tuple,
    ),
    kids=[
        StringFileInfo([
            StringTable("040904B0", [
                StringStruct("CompanyName", "vanilla1108"),
                StringStruct("FileDescription", "pptx2md GUI — PPT/PPTX to Markdown converter"),
                StringStruct("FileVersion", APP_VERSION),
                StringStruct("InternalName", "pptx2md-gui"),
                StringStruct("LegalCopyright", "Apache License 2.0"),
                StringStruct("OriginalFilename", f"pptx2md-gui-{APP_VERSION}.exe"),
                StringStruct("ProductName", "pptx2md-gui"),
                StringStruct("ProductVersion", APP_VERSION),
            ]),
        ]),
        VarFileInfo([VarStruct("Translation", [0x0409, 1200])]),
    ],
)

APP_NAME = f"pptx2md-gui-{APP_VERSION}"

# ---------------------------------------------------------------------------
# Hidden imports
# ---------------------------------------------------------------------------

# pydantic v2 relies on compiled Rust core + annotated_types
hiddenimports = [
    "pydantic",
    "pydantic.deprecated",
    "pydantic.deprecated.decorator",
    "pydantic_core",
    "pydantic_core._pydantic_core",
    "annotated_types",
]

# All submodules of our own packages (including lazy-imported ppt_legacy)
hiddenimports += collect_submodules("pptx2md")
hiddenimports += collect_submodules("pptx2md_gui")
hiddenimports += collect_submodules("pptx2md.ppt_legacy")

# dwml (OMML to LaTeX)
hiddenimports += collect_submodules("dwml")

# scipy.optimize.curve_fit used in multi_column.py
hiddenimports += [
    "scipy.optimize",
    "scipy.optimize._minpack_py",
    "scipy.special",
    "scipy.special._ufuncs",
]

# rapidfuzz native extensions
hiddenimports += collect_submodules("rapidfuzz")

# customtkinter internals
hiddenimports += collect_submodules("customtkinter")

# pywin32 / COM (Windows only): .ppt conversion + WMF COM fallback
if sys.platform == "win32":
    hiddenimports += collect_submodules("win32com")
    hiddenimports += collect_submodules("win32comext")
    hiddenimports += ["pythoncom", "pywintypes", "win32timezone"]

# tkinterdnd2
hiddenimports += ["tkinterdnd2", "tkinterdnd2.TkinterDnD"]

# CTkToolTip
hiddenimports += collect_submodules("CTkToolTip")

# tqdm (used for CLI progress bars)
hiddenimports += ["tqdm", "tqdm.auto"]

# lxml (used internally by python-pptx)
hiddenimports += collect_submodules("lxml")

# PIL / Pillow
hiddenimports += collect_submodules("PIL")

# numpy
hiddenimports += ["numpy", "numpy.core", "numpy.core._multiarray_umath"]

# ---------------------------------------------------------------------------
# Data files
# ---------------------------------------------------------------------------

datas = []
binaries = []


def _iter_unique_prefix_roots():
    """返回当前解释器相关的唯一前缀目录。"""
    seen = set()
    for raw_prefix in (sys.prefix, sys.base_prefix):
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


def _is_within_prefix_roots(path: Path) -> bool:
    """判断文件是否位于当前解释器环境目录下。"""
    try:
        resolved = path.resolve()
    except OSError:
        resolved = Path(os.path.abspath(path))

    for prefix_root in _iter_unique_prefix_roots():
        try:
            resolved.relative_to(prefix_root)
            return True
        except ValueError:
            continue
    return False


def _append_binary(source_path, destdir="."):
    """向 binaries 中追加单个动态库，自动忽略缺失项和重复项。"""
    if not source_path:
        return

    source = Path(source_path)
    if not source.is_file():
        return

    normalized_source = os.path.normcase(str(source.resolve()))
    for existing_source, existing_dest in binaries:
        try:
            existing_norm = os.path.normcase(str(Path(existing_source).resolve()))
        except OSError:
            existing_norm = os.path.normcase(os.path.abspath(existing_source))
        if existing_dest == destdir and existing_norm == normalized_source:
            return

    binaries.append((str(source), destdir))


def _append_env_scoped_binary(source_path, destdir="."):
    """仅收集当前解释器环境内的动态库，避免 PATH 上的同名外部 DLL 混入。"""
    if not source_path:
        return

    source = Path(source_path)
    if not source.is_file():
        return

    if not _is_within_prefix_roots(source):
        print(f"Skipping external binary outside current env: {source}")
        return

    _append_binary(source, destdir)

# customtkinter theme JSON files + assets
datas += collect_data_files("customtkinter")

# tkinterdnd2 native libraries (.dll, .tcl)
datas += collect_data_files("tkinterdnd2")

# CTkToolTip
datas += collect_data_files("CTkToolTip")

# python-pptx templates (default.pptx, theme.xml, etc.)
datas += collect_data_files("pptx")

# pywin32 native DLLs (pythoncom/pywintypes)
if sys.platform == "win32":
    binaries += collect_dynamic_libs("pywin32_system32")

# Tcl/Tk runtime:
# 1. 让 PyInstaller 的 tcltk_info 参与收集；
# 2. 但只接受当前解释器环境内的 DLL，避免 Graphviz 等 PATH 前置软件把
#    旧版 tcl86t.dll / tk86t.dll 混入，造成 _tcl_data 与 DLL 版本不一致。
_append_env_scoped_binary(getattr(tcltk_info, "tcl_shared_library", None))
_append_env_scoped_binary(getattr(tcltk_info, "tk_shared_library", None))

if sys.platform == "win32":
    # Conda 会把部分标准库扩展的底层依赖放到 Library/bin，PyInstaller 在
    # 该环境下不会稳定收集它们，导致 _tkinter / _ctypes / _bz2 / _decimal
    # 在启动阶段因缺少 DLL 直接崩溃。
    for _base_prefix in {Path(sys.prefix), Path(sys.base_prefix)}:
        _library_bin = _base_prefix / "Library" / "bin"
        for _dll_name in (
            "tcl86t.dll",
            "tk86t.dll",
            "ffi.dll",
            "libbz2.dll",
            "libmpdec-4.dll",
        ):
            _append_binary(_library_bin / _dll_name)

# ---------------------------------------------------------------------------
# Analysis
# ---------------------------------------------------------------------------

a = Analysis(
    [os.path.join(PROJECT_ROOT, "pptx2md_gui", "__main__.py")],
    pathex=[PROJECT_ROOT],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # === Dev / test only ===
        "wand", "pytest", "_pytest", "yapf", "isort", "pycln",
        "sphinx",
        # === GUI frameworks we don't use ===
        "PyQt5", "PyQt6", "PySide2", "PySide6", "qtpy", "sip",
        # === Heavy science / data packages not needed ===
        "matplotlib", "pandas", "sklearn", "scikit-learn",
        "bokeh", "plotly", "altair", "panel", "pyviz_comms",
        "statsmodels", "patsy", "xarray",
        "astropy", "astropy_iers_data",
        "skimage", "scikit-image", "imageio",
        "h5py", "tables", "pytables",
        "numba", "llvmlite", "numexpr",
        "dask", "distributed", "fsspec",
        "pyarrow", "openpyxl",
        "sqlalchemy",
        # === Jupyter / notebook ===
        "IPython", "jupyter", "notebook", "nbformat", "nbconvert",
        "ipykernel", "ipywidgets", "traitlets",
        # === Web / network ===
        "flask", "werkzeug", "jinja2",
        "tornado", "aiohttp", "yarl", "multidict",
        "botocore", "boto3", "s3transfer",
        # === Serialization / compression not needed ===
        "zmq", "pyzmq", "msgpack", "ujson",
        "lz4", "zstandard", "blosc", "brotlicffi",
        "ruamel", "ruamel.yaml",
        # === Crypto / security not needed ===
        "cryptography", "bcrypt", "argon2",
        # === Other unused packages ===
        "mpi4py", "psutil", "greenlet",
        "docutils", "markdown",
        "jsonschema", "jsonschema_specifications",
        "xyzservices", "intake",
        "mypy", "pycosat", "winloop",
        "gmpy2", "pyreadline3",
        "lmdb", "pycares",
        "pywt", "contourpy",
        "click",
    ],
    noarchive=False,
    cipher=block_cipher,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

# ---------------------------------------------------------------------------
# EXE
# ---------------------------------------------------------------------------

if IS_ONEFILE:
    # onefile: binaries/datas directly packed into the executable.
    exe = EXE(
        pyz,
        a.scripts,
        a.binaries,
        a.datas,
        [],
        name=APP_NAME,
        version=_version_info,
        debug=False,
        bootloader_ignore_signals=False,
        strip=False,
        upx=False,  # Disable UPX to avoid AV false positives and WinError 5
        console=False,  # GUI application, no console window
        disable_windowed_traceback=False,
        argv_emulation=False,
        target_arch=None,
        codesign_identity=None,
        entitlements_file=None,
    )
else:
    # onedir: more stable and easier to inspect.
    exe = EXE(
        pyz,
        a.scripts,
        [],
        exclude_binaries=True,
        name=APP_NAME,
        version=_version_info,
        debug=False,
        bootloader_ignore_signals=False,
        strip=False,
        upx=False,  # Disable UPX to avoid AV false positives and WinError 5
        console=False,  # GUI application, no console window
        disable_windowed_traceback=False,
        argv_emulation=False,
        target_arch=None,
        codesign_identity=None,
        entitlements_file=None,
    )

    coll = COLLECT(
        exe,
        a.binaries,
        a.datas,
        strip=False,
        upx=False,
        upx_exclude=[],
        name=APP_NAME,
    )
