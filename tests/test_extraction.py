"""内容提取对比测试：断言已知的 PPTX 内容正确出现在输出中。

所有断言值来源于 test_output/3 深度学习概览2025.md 参考输出。
使用 session-scoped dl_output_text fixture 共享转换结果。
"""

import re

import pytest

from tests.conftest import split_by_slides


# ---------------------------------------------------------------------------
# 标题测试
# ---------------------------------------------------------------------------

@pytest.mark.slow
def test_first_title_is_h1(dl_output_text):
    """Slide 1 的标题应为 H1：# 深度学习和大模型基础。"""
    slides = split_by_slides(dl_output_text)
    assert "# 深度学习和大模型基础" in slides[1]


@pytest.mark.slow
def test_section_titles_are_h2(dl_output_text):
    """Slide 5/6 应包含 H2 标题。"""
    slides = split_by_slides(dl_output_text)
    assert "## 传统机器学习" in slides[5]
    assert "## 传统机器学习与深度学习" in slides[6]


# ---------------------------------------------------------------------------
# 列表测试
# ---------------------------------------------------------------------------

@pytest.mark.slow
def test_unordered_list_slide2(dl_output_text):
    """Slide 2 应包含无序列表（>= 3 项）和特定内容。"""
    slides = split_by_slides(dl_output_text)
    slide2 = slides[2]

    bullet_items = re.findall(r"^\* .+", slide2, re.MULTILINE)
    assert len(bullet_items) >= 3, (
        f"Expected >= 3 bullet items, got {len(bullet_items)}"
    )
    assert "在当今数字化时代" in slide2


@pytest.mark.slow
def test_ordered_list_slide4(dl_output_text):
    """Slide 4 应包含有序列表和加粗的感知机。"""
    slides = split_by_slides(dl_output_text)
    slide4 = slides[4]

    assert re.search(r"^1\. ", slide4, re.MULTILINE), (
        "Slide 4 should contain ordered list starting with '1. '"
    )
    assert "__感知机__" in slide4, "Slide 4 should contain bold '感知机'"


@pytest.mark.slow
def test_nested_list_slide3(dl_output_text):
    """Slide 3 应包含段落文本 + 缩进无序列表。"""
    slides = split_by_slides(dl_output_text)
    slide3 = slides[3]

    # 段落文本
    assert "学完本课程后" in slide3
    # 缩进列表项
    assert re.search(r"^ {2}\* 掌握深度学习基础", slide3, re.MULTILINE), (
        "Slide 3 should contain indented list item '  * 掌握深度学习基础'"
    )


# ---------------------------------------------------------------------------
# 表格测试
# ---------------------------------------------------------------------------

@pytest.mark.slow
def test_table_slide6(dl_output_text):
    """Slide 6 应包含 Markdown 表格。"""
    slides = split_by_slides(dl_output_text)
    slide6 = slides[6]

    # 表头
    assert "__传统机器学习__" in slide6 or "传统机器学习" in slide6
    # 对齐标记
    assert ":-:" in slide6, "Table should contain center alignment markers"
    # 内容
    assert "端到端" in slide6


# ---------------------------------------------------------------------------
# 图片测试
# ---------------------------------------------------------------------------

@pytest.mark.slow
def test_image_slide7(dl_output_text):
    """Slide 7 应包含图片引用。"""
    slides = split_by_slides(dl_output_text)
    slide7 = slides[7]

    assert re.search(r"!\[.*?\]\(.*?\)", slide7), (
        "Slide 7 should contain an image reference"
    )


# ---------------------------------------------------------------------------
# 数学公式测试
# ---------------------------------------------------------------------------

@pytest.mark.slow
def test_math_inline_slide9(dl_output_text):
    """Slide 9 应包含行内公式片段。"""
    slides = split_by_slides(dl_output_text)
    slide9 = slides[9]

    assert "$X=[x_{0}" in slide9 or "$X=[x_{" in slide9, (
        "Slide 9 should contain inline math formula X=[x_{0}..."
    )


@pytest.mark.slow
def test_math_block_slide9(dl_output_text):
    """Slide 9 应包含块级公式。"""
    slides = split_by_slides(dl_output_text)
    slide9 = slides[9]

    assert "Ax+By" in slide9, "Slide 9 should contain 'Ax+By' formula"
    assert "$W^{T}X+b=0$" in slide9 or "$W^{T}" in slide9, (
        "Slide 9 should contain W^{T}X+b=0 formula"
    )


@pytest.mark.slow
def test_math_formula_count(dl_output_text):
    """全文 $ 符号应 >= 400（大量 LaTeX 公式）。"""
    dollar_count = dl_output_text.count("$")
    assert dollar_count >= 400, (
        f"Expected >= 400 $ signs, got {dollar_count}"
    )


# ---------------------------------------------------------------------------
# 中文文本测试
# ---------------------------------------------------------------------------

@pytest.mark.slow
def test_chinese_text_slide9(dl_output_text):
    """Slide 9 应包含关键中文术语。"""
    slides = split_by_slides(dl_output_text)
    slide9 = slides[9]

    for term in ("输入向量", "权值", "激活函数", "分割平面", "分割超平面"):
        assert term in slide9, f"Slide 9 should contain '{term}'"


# ---------------------------------------------------------------------------
# 结构连续性测试
# ---------------------------------------------------------------------------

@pytest.mark.slow
def test_slide_number_comments(dl_output_text):
    """slide 注释应从 1 到 141 连续。"""
    slides = split_by_slides(dl_output_text)
    expected = set(range(1, 142))
    actual = set(slides.keys())
    assert actual == expected, (
        f"Missing slides: {expected - actual}, "
        f"Extra slides: {actual - expected}"
    )
