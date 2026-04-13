from pathlib import Path

from pptx2md.outputter import MarkdownFormatter
from pptx2md.types import ConversionConfig, TextRun, TextStyle


def _make_markdown_formatter(tmp_path: Path) -> MarkdownFormatter:
    config = ConversionConfig(
        pptx_path=tmp_path / "dummy.pptx",
        output_path=tmp_path / "result.md",
        image_dir=tmp_path / "img",
        disable_color=True,
        compress_blank_lines=False,
    )
    return MarkdownFormatter(config)


def test_markdown_formatter_merges_adjacent_strong_runs(tmp_path):
    formatter = _make_markdown_formatter(tmp_path)
    runs = [
        TextRun(text="1.1                     ", style=TextStyle(is_strong=True)),
        TextRun(text="计算机网络", style=TextStyle(is_strong=True)),
        TextRun(text="在信息时代中的", style=TextStyle(is_strong=True)),
        TextRun(text="作用", style=TextStyle(is_strong=True)),
    ]

    try:
        formatted = formatter.get_formatted_runs(runs)
    finally:
        formatter.close()

    assert formatted == "**1\\.1                     计算机网络在信息时代中的作用**"
    assert "__" not in formatted


def test_markdown_formatter_keeps_trailing_spaces_outside_strong_markers(tmp_path):
    formatter = _make_markdown_formatter(tmp_path)
    runs = [
        TextRun(text="重点  ", style=TextStyle(is_strong=True)),
        TextRun(text="说明", style=TextStyle()),
    ]

    try:
        formatted = formatter.get_formatted_runs(runs)
    finally:
        formatter.close()

    assert formatted == "**重点**  说明"


def test_markdown_formatter_merges_effectively_same_runs_when_color_disabled(tmp_path):
    formatter = _make_markdown_formatter(tmp_path)
    runs = [
        TextRun(text="发展最快的并起到核心作用的是", style=TextStyle(is_strong=True, color_rgb=(0, 0, 0))),
        TextRun(text="计算机网络。", style=TextStyle(is_strong=True, color_rgb=(255, 0, 0))),
    ]

    try:
        formatted = formatter.get_formatted_runs(runs)
    finally:
        formatter.close()

    assert formatted == "**发展最快的并起到核心作用的是计算机网络。**"
    assert "****" not in formatted
