"""结构冒烟测试：验证转换管道输出的完整性和基本结构。

所有测试使用 session-scoped fixture 共享转换结果，避免重复转换。
"""

import re

import pytest

from tests.conftest import (
    SAMPLE_PPTX,
    SAMPLE_PPTX_DL,
    SMOKE_CONFIG_OVERRIDES,
    _run_conversion,
    split_by_slides,
)


# ---------------------------------------------------------------------------
# 深度学习概览样本冒烟测试
# ---------------------------------------------------------------------------

@pytest.mark.slow
def test_dl_output_nonempty(dl_output_text):
    """输出文本非空。"""
    assert len(dl_output_text.strip()) > 0


@pytest.mark.slow
def test_dl_output_min_length(dl_output_text):
    """输出至少 1000 行。"""
    lines = dl_output_text.splitlines()
    assert len(lines) >= 1000, f"Expected >= 1000 lines, got {len(lines)}"


@pytest.mark.slow
def test_dl_slide_count(dl_output_text):
    """深度学习概览应有 141 页 slide。"""
    slides = split_by_slides(dl_output_text)
    assert len(slides) == 141, f"Expected 141 slides, got {len(slides)}"


@pytest.mark.slow
def test_dl_has_titles(dl_output_text):
    """输出应包含足够多的标题（#{1,3}）。"""
    title_count = len(re.findall(r"^#{1,3} ", dl_output_text, re.MULTILINE))
    assert title_count >= 50, f"Expected >= 50 titles, got {title_count}"


@pytest.mark.slow
def test_dl_has_slide_separators(dl_output_text):
    """幻灯片分隔符 --- 应为 140 个（141 页间有 140 个分隔）。"""
    sep_count = len(re.findall(r"^---$", dl_output_text, re.MULTILINE))
    assert sep_count == 140, f"Expected 140 separators, got {sep_count}"


@pytest.mark.slow
def test_dl_has_images(dl_smoke_result):
    """输出应包含图片引用和图片文件。"""
    text = dl_smoke_result.text
    image_dir = dl_smoke_result.image_dir

    image_refs = re.findall(r"!\[.*?\]\(.*?\)", text)
    assert len(image_refs) >= 50, (
        f"Expected >= 50 image references, got {len(image_refs)}"
    )

    if image_dir.exists():
        image_files = list(image_dir.glob("*"))
        assert len(image_files) >= 50, (
            f"Expected >= 50 image files, got {len(image_files)}"
        )


# ---------------------------------------------------------------------------
# 人工智能前沿应用场景样本冒烟测试
# ---------------------------------------------------------------------------

@pytest.mark.slow
def test_ai_output_nonempty(ai_output_text):
    """第二样本输出非空且页数 >= 10。"""
    assert len(ai_output_text.strip()) > 0
    slides = split_by_slides(ai_output_text)
    assert len(slides) >= 10, f"Expected >= 10 slides, got {len(slides)}"


# ---------------------------------------------------------------------------
# 格式变体测试
# ---------------------------------------------------------------------------

@pytest.mark.slow
def test_wiki_format_converts(workspace_tmp_path):
    """Wiki 格式输出应包含 ! 标题标记。"""
    result = _run_conversion(
        SAMPLE_PPTX_DL,
        workspace_tmp_path / "wiki",
        is_wiki=True,
        page=1,
        **{k: v for k, v in SMOKE_CONFIG_OVERRIDES.items()
           if k not in ("enable_slides",)},
    )
    assert "!" in result.text, "Wiki output should contain ! heading markers"


@pytest.mark.slow
def test_quarto_format_converts(workspace_tmp_path):
    """Quarto 格式输出应包含 YAML 头和 revealjs。"""
    result = _run_conversion(
        SAMPLE_PPTX_DL,
        workspace_tmp_path / "quarto",
        is_qmd=True,
        page=1,
        disable_color=True,
        disable_escaping=True,
        disable_notes=True,
        min_block_size=3,
    )
    assert "---" in result.text, "Quarto output should contain YAML front matter"
    assert "revealjs" in result.text, "Quarto output should specify revealjs format"


# ---------------------------------------------------------------------------
# 确定性测试
# ---------------------------------------------------------------------------

@pytest.mark.slow
def test_determinism(workspace_tmp_path):
    """同配置转换两次，输出应完全相同。"""
    overrides = {**SMOKE_CONFIG_OVERRIDES, "page": 1}

    result_a = _run_conversion(
        SAMPLE_PPTX_DL,
        workspace_tmp_path / "det_a",
        **overrides,
    )
    result_b = _run_conversion(
        SAMPLE_PPTX_DL,
        workspace_tmp_path / "det_b",
        **overrides,
    )
    assert result_a.text == result_b.text, "Same config should produce identical output"
