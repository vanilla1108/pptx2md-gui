# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec file for pptx2md-gui portable build."""

import os
import re
import sys
import tomllib
from pathlib import Path

from PyInstaller.utils.hooks import collect_data_files, collect_dynamic_libs, collect_submodules
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

# Parse "0.1.0b3" -> (0, 1, 0, 3) for Windows version resource
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
