"""测试共享 fixture 和工具函数。

提供：
- 样本 PPTX 路径常量和 fixture
- Session-scoped 转换结果 fixture（每个样本只转换一次）
- 文本分割工具（按 slide 注释分割）
"""

import os
import re
import tempfile
from pathlib import Path
from typing import NamedTuple
from uuid import uuid4

import pytest

import pptx2md.parser as _parser_module
from pptx2md.entry import convert
from pptx2md.types import ConversionConfig

# ---------------------------------------------------------------------------
# 路径常量
# ---------------------------------------------------------------------------

PROJECT_ROOT = Path(__file__).resolve().parent.parent
SAMPLE_PPTX = PROJECT_ROOT / "test_pptx" / "6.人工智能前沿应用场景2025.pptx"
SAMPLE_PPTX_DL = PROJECT_ROOT / "test_pptx" / "3 深度学习概览2025.pptx"

# ---------------------------------------------------------------------------
# 临时目录策略（适配受限运行环境）
# ---------------------------------------------------------------------------
#
# pytest 的 tmp_path/tmpdir 由其插件提供（见 pyproject.toml 的 --basetemp=...）。
# 这里额外把标准库 tempfile 也指向项目内可写目录，避免使用系统 Temp 导致权限/清理问题。
#
_TEMP_BASE = PROJECT_ROOT / "tmp_test_artifacts" / "std_temp"
_TEMP_BASE.mkdir(parents=True, exist_ok=True)
os.environ.setdefault("TMP", str(_TEMP_BASE))
os.environ.setdefault("TEMP", str(_TEMP_BASE))
tempfile.tempdir = str(_TEMP_BASE)


# ---------------------------------------------------------------------------
# 转换结果类型
# ---------------------------------------------------------------------------

class ConversionResult(NamedTuple):
    text: str
    image_dir: Path


# ---------------------------------------------------------------------------
# 冒烟测试共用配置
# ---------------------------------------------------------------------------

SMOKE_CONFIG_OVERRIDES = dict(
    disable_color=True,
    disable_escaping=True,
    disable_notes=True,
    enable_slides=True,
    disable_slide_number=False,
    min_block_size=3,
    keep_similar_titles=True,
    compress_blank_lines=True,
)


# ---------------------------------------------------------------------------
# 基础 fixture（function-scoped，test_math_formula.py 使用）
# ---------------------------------------------------------------------------

@pytest.fixture
def sample_pptx() -> Path:
    """Path to the shared test PPTX file (人工智能前沿应用场景)."""
    assert SAMPLE_PPTX.exists(), f"Test sample not found: {SAMPLE_PPTX}"
    return SAMPLE_PPTX


@pytest.fixture
def sample_pptx_dl() -> Path:
    """Path to the deep learning test PPTX file (深度学习概览)."""
    assert SAMPLE_PPTX_DL.exists(), f"Test sample not found: {SAMPLE_PPTX_DL}"
    return SAMPLE_PPTX_DL


@pytest.fixture
def workspace_tmp_path() -> Path:
    """在项目目录下创建一个不会自动清理的临时目录。

    说明：当前运行环境对删除文件/目录有限制，pytest 内置 tmp_path
    夹具在清理阶段会触发 PermissionError，因此这里提供一个替代夹具。
    """
    base = PROJECT_ROOT / "tmp_test_artifacts"
    base.mkdir(parents=True, exist_ok=True)

    run_id = f"{os.getpid()}_{uuid4().hex[:8]}"
    path = base / f"run_{run_id}"
    path.mkdir(parents=True, exist_ok=False)
    return path


@pytest.fixture
def tmp_path(workspace_tmp_path) -> Path:
    """兼容 pytest 的 tmp_path fixture 名称，但不依赖 tmpdir 插件也不自动清理。"""
    return workspace_tmp_path


# ---------------------------------------------------------------------------
# 内部工具
# ---------------------------------------------------------------------------

def _reset_picture_counter():
    """重置 parser.picture_count 全局计数器。"""
    _parser_module.picture_count = 0


def _run_conversion(pptx: Path, output_dir: Path, **overrides) -> ConversionResult:
    """执行单次 PPTX 转换并返回结果。

    参数:
        pptx: 输入 PPTX 路径。
        output_dir: 输出目录（自动创建 result.md 和 img/ 子目录）。
        **overrides: 传给 ConversionConfig 的额外参数。

    返回:
        ConversionResult(text, image_dir)。
    """
    output_dir.mkdir(parents=True, exist_ok=True)

    ext = ".md"
    if overrides.get("is_wiki"):
        ext = ".tid"
    elif overrides.get("is_qmd"):
        ext = ".qmd"

    output_path = output_dir / f"result{ext}"
    image_dir = output_dir / "img"

    _reset_picture_counter()
    config = ConversionConfig(
        pptx_path=pptx,
        output_path=output_path,
        image_dir=image_dir,
        **overrides,
    )
    convert(config, disable_tqdm=True)

    text = output_path.read_text(encoding="utf-8")
    return ConversionResult(text=text, image_dir=image_dir)


# ---------------------------------------------------------------------------
# Session-scoped fixture
# ---------------------------------------------------------------------------

@pytest.fixture(scope="session")
def session_tmp_dir() -> Path:
    """会话级临时目录（不自动清理）。"""
    base = PROJECT_ROOT / "tmp_test_artifacts"
    base.mkdir(parents=True, exist_ok=True)
    path = base / f"session_{os.getpid()}_{uuid4().hex[:8]}"
    path.mkdir(parents=True, exist_ok=False)
    return path


@pytest.fixture(scope="session")
def dl_smoke_result(session_tmp_dir) -> ConversionResult:
    """深度学习概览样本的转换结果（session 共享）。"""
    assert SAMPLE_PPTX_DL.exists(), f"Test sample not found: {SAMPLE_PPTX_DL}"
    return _run_conversion(
        SAMPLE_PPTX_DL,
        session_tmp_dir / "dl_smoke",
        **SMOKE_CONFIG_OVERRIDES,
    )


@pytest.fixture(scope="session")
def ai_smoke_result(session_tmp_dir) -> ConversionResult:
    """人工智能前沿应用场景样本的转换结果（session 共享）。"""
    assert SAMPLE_PPTX.exists(), f"Test sample not found: {SAMPLE_PPTX}"
    return _run_conversion(
        SAMPLE_PPTX,
        session_tmp_dir / "ai_smoke",
        **SMOKE_CONFIG_OVERRIDES,
    )


@pytest.fixture(scope="session")
def dl_output_text(dl_smoke_result) -> str:
    """深度学习概览输出文本（快捷访问）。"""
    return dl_smoke_result.text


@pytest.fixture(scope="session")
def ai_output_text(ai_smoke_result) -> str:
    """人工智能前沿应用场景输出文本（快捷访问）。"""
    return ai_smoke_result.text


# ---------------------------------------------------------------------------
# 文本工具
# ---------------------------------------------------------------------------

def split_by_slides(text: str) -> dict[int, str]:
    """按 <!-- slide: N --> 注释分割文本，返回 {页码: 该页内容}。"""
    pattern = re.compile(r"<!-- slide: (\d+) -->")
    result: dict[int, str] = {}
    parts = pattern.split(text)

    # parts: [前缀, "1", slide1内容, "2", slide2内容, ...]
    for i in range(1, len(parts), 2):
        slide_num = int(parts[i])
        content = parts[i + 1] if i + 1 < len(parts) else ""
        result[slide_num] = content

    return result
