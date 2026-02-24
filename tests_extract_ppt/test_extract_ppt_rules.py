"""extract_ppt 脚本级规则测试（无 COM 依赖）。"""

from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace
from uuid import uuid4

from ppt2md_script import extract_ppt as ep
from ppt2md_script import renderer_markdown as rm


class _FakeBullet:
    def __init__(self, visible=False, bullet_type=0, start_value=1):
        self.Visible = visible
        self.Type = bullet_type
        self.StartValue = start_value


class _FakeParagraphFormat:
    def __init__(self, bullet_visible=False, bullet_type=0, start_value=1):
        self.Bullet = _FakeBullet(
            visible=bullet_visible,
            bullet_type=bullet_type,
            start_value=start_value,
        )


class _FakeParagraph:
    def __init__(self, text, indent_level=1, bullet_visible=False, bullet_type=0, start_value=1):
        self.Text = text
        self.IndentLevel = indent_level
        self.ParagraphFormat = _FakeParagraphFormat(
            bullet_visible=bullet_visible,
            bullet_type=bullet_type,
            start_value=start_value,
        )


class _FakeTextRange:
    def __init__(self, paragraphs):
        self._paragraphs = list(paragraphs)
        self.Text = "\n".join([str(p.Text) for p in self._paragraphs])

    def Paragraphs(self, index=None, _count=None):
        if index is None:
            return SimpleNamespace(Count=len(self._paragraphs))
        return self._paragraphs[int(index) - 1]


class _FakeTextFrame:
    def __init__(self, paragraphs):
        self.HasText = bool(paragraphs)
        self.TextRange = _FakeTextRange(paragraphs)


class _FakeShape:
    def __init__(self, shape_type=1, shape_id=1, paragraphs=None, alt_text="", export_ok=True):
        self.Type = shape_type
        self.Id = shape_id
        self.HasTable = False
        self.AlternativeText = alt_text
        self._export_ok = bool(export_ok)
        if paragraphs is None:
            self.HasTextFrame = False
            self.TextFrame = None
        else:
            self.HasTextFrame = True
            self.TextFrame = _FakeTextFrame(paragraphs)

    def Export(self, path, _fmt):
        if not self._export_ok:
            raise RuntimeError("mock export failed")
        Path(path).write_bytes(b"fake-png-data")


def _make_workspace_dir() -> Path:
    root = Path("tmp_test_artifacts") / "tests_extract_ppt"
    root.mkdir(parents=True, exist_ok=True)
    path = root / f"run_{uuid4().hex[:8]}"
    path.mkdir(parents=True, exist_ok=False)
    return path


def test_manual_ordered_prefix_to_markdown_numbered_line():
    shape = _FakeShape(
        paragraphs=[
            _FakeParagraph("1、 处理器模式", indent_level=1, bullet_visible=False),
        ]
    )

    lines = ep.extract_text_from_shape(shape)

    assert lines == ["1. 处理器模式"]


def test_manual_ordered_block_nested_child_auto_unordered_item():
    shape = _FakeShape(
        paragraphs=[
            _FakeParagraph("1、 一级条目", indent_level=1, bullet_visible=False),
            _FakeParagraph("2、 二级条目", indent_level=1, bullet_visible=False),
            _FakeParagraph("子项说明", indent_level=2, bullet_visible=False),
        ]
    )

    lines = ep.extract_text_from_shape(shape)

    assert lines[0] == "1. 一级条目"
    assert lines[1] == "2. 二级条目"
    assert lines[2] == "  - 子项说明"


def test_renderer_comment_output_format():
    line = rm.md_comment("slide: 2 | path: S2")
    escaped = rm.md_comment("a --> b\nc")

    assert line == "<!-- slide: 2 | path: S2 -->\n"
    assert escaped == "<!-- a --＞ b c -->\n"


def test_image_export_markdown_uses_relative_path():
    workdir = _make_workspace_dir()
    output_path = workdir / "result.md"

    image_ctx = ep._build_image_extract_context(str(output_path), extract_images=True, image_dir="img")
    shape = _FakeShape(shape_type=13, shape_id=7, alt_text="封面图", paragraphs=None, export_ok=True)

    md = ep._export_shape_image_markdown(shape, image_ctx=image_ctx, image_loc="S1/R1/SH7")

    assert md.startswith("![图片: 封面图](img/")
    assert ".png)" in md
    assert "\\" not in md


def test_image_export_fallback_to_placeholder_when_disabled():
    shape = _FakeShape(shape_type=13, shape_id=9, alt_text="方括号]测试", paragraphs=None, export_ok=True)
    image_ctx = {"enabled": False, "dir": None, "md_dir": "", "counter": 0}

    md = ep._export_shape_image_markdown(shape, image_ctx=image_ctx, image_loc="S1/R1/SH9")

    assert md == "![图片: 方括号\\]测试]"
