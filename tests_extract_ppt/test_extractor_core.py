"""extractor_core 模块测试。"""

from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace
from uuid import uuid4

from ppt2md_script import extractor_core as core


class _Shape:
    def __init__(self, shape_type=1, prog_id=None):
        self.Type = shape_type
        if prog_id is not None:
            self.OLEFormat = SimpleNamespace(ProgID=prog_id)


class _Paragraph:
    def __init__(self, text, indent_level=1):
        self.Text = text
        self.IndentLevel = indent_level


class _TextRange:
    def __init__(self, paragraphs):
        self._paragraphs = list(paragraphs)

    def Paragraphs(self, index=None, _count=None):
        if index is None:
            return SimpleNamespace(Count=len(self._paragraphs))
        return self._paragraphs[int(index) - 1]


class _TextFrame:
    def __init__(self, paragraphs):
        self.HasText = bool(paragraphs)
        self.TextRange = _TextRange(paragraphs)


class _TextShape:
    def __init__(self, shape_type=1, has_table=False, paragraphs=None):
        self.Type = shape_type
        self.HasTable = bool(has_table)
        self.HasTextFrame = paragraphs is not None
        self.TextFrame = _TextFrame(paragraphs or []) if paragraphs is not None else None


def _make_workspace_dir() -> Path:
    root = Path("tmp_test_artifacts") / "tests_extract_ppt_core"
    root.mkdir(parents=True, exist_ok=True)
    path = root / f"run_{uuid4().hex[:8]}"
    path.mkdir(parents=True, exist_ok=False)
    return path


def test_split_row_shapes_for_embedding():
    normal = _Shape(shape_type=1)
    embedded_ppt = _Shape(shape_type=7, prog_id="PowerPoint.Show.12")
    embedded_excel = _Shape(shape_type=7, prog_id="Excel.Sheet.12")

    normals, embedded_ppts, embedded_objects = core.split_row_shapes_for_embedding(
        [normal, embedded_ppt, embedded_excel]
    )

    assert normals == [normal]
    assert embedded_ppts == [embedded_ppt]
    assert embedded_objects == ["Excel.Sheet.12"]


def test_is_list_block_core_detects_mixed_indent():
    shape = _TextShape(paragraphs=[_Paragraph("A", indent_level=1), _Paragraph("B", indent_level=2)])

    assert core.is_list_block_core(shape) is True


def test_get_single_line_plain_text_core_returns_text_for_single_para_non_list():
    shape = _TextShape(paragraphs=[_Paragraph(" 单行标题 ", indent_level=1)])

    text = core.get_single_line_plain_text_core(shape, is_list_block_fn=core.is_list_block_core)

    assert text == "单行标题"


def test_get_single_line_plain_text_core_returns_none_for_list_block():
    shape = _TextShape(paragraphs=[_Paragraph("A", indent_level=1), _Paragraph("B", indent_level=2)])

    text = core.get_single_line_plain_text_core(shape, is_list_block_fn=core.is_list_block_core)

    assert text is None


def test_split_manual_ordered_prefix_core():
    assert core.split_manual_ordered_prefix_core(" 12、 指令流水线 ") == (12, "指令流水线")
    assert core.split_manual_ordered_prefix_core("普通文本") is None


def test_strip_bullet_like_prefix_core():
    assert core.strip_bullet_like_prefix_core("► 分支预测") == "分支预测"
    assert core.strip_bullet_like_prefix_core("正文") is None


def test_looks_like_brief_list_item_core():
    assert core.looks_like_brief_list_item_core("短句子") is True
    assert core.looks_like_brief_list_item_core("这是一个完整句子。") is False


def test_escape_md_text_line_core():
    assert core.escape_md_text_line_core("# 标题") == r"\# 标题"
    assert core.escape_md_text_line_core("1. 项") == r"\1. 项"


def test_escape_md_table_cell_core():
    assert core.escape_md_table_cell_core("a|b\nc") == r"a\|b<br>c"


def test_get_unique_output_path_core():
    existing = {r"D:\out\demo.md", r"D:\out\demo_1.md"}
    path = core.get_unique_output_path_core(r"D:\out\demo.md", path_exists_fn=lambda p: p in existing)
    assert path == r"D:\out\demo_2.md"


def test_read_shape_alt_text_core():
    shape = SimpleNamespace(AlternativeText=" 封面\r\n图 ")
    assert core.read_shape_alt_text_core(shape) == "封面 图"


def test_safe_shape_id_core():
    assert core.safe_shape_id_core(SimpleNamespace(Id="15")) == 15
    assert core.safe_shape_id_core(SimpleNamespace()) is None


def test_detect_slide_title_core_prefers_title_placeholder():
    title_shape = SimpleNamespace(Id=3, mock_text="主标题")

    class _Shapes(list):
        pass

    shapes = _Shapes([])
    shapes.Title = title_shape
    slide = SimpleNamespace(Shapes=shapes)

    info = core.detect_slide_title_core(
        slide,
        safe_shape_id_fn=core.safe_shape_id_core,
        first_paragraph_text_fn=lambda shape: getattr(shape, "mock_text", None),
        is_title_candidate_shape_fn=lambda _shape: False,
    )

    assert info == {"title": "主标题", "shape_id": 3}


def test_extract_title_shape_extra_lines_core_filters_title_dup():
    slide = SimpleNamespace()
    title_info = {"shape_id": 9, "title": "标题"}
    shape = SimpleNamespace(Id=9)

    lines = core.extract_title_shape_extra_lines_core(
        slide,
        title_info,
        find_shape_by_id_fn=lambda _slide, _sid: shape,
        extract_text_from_shape_fn=lambda _shape, skip_first_para_text=None: [skip_first_para_text, "副标题"],
    )

    assert lines == ["副标题"]


def test_build_title_render_context_core():
    slide = SimpleNamespace()

    ctx = core.build_title_render_context_core(
        slide,
        fallback_title="默认标题",
        detect_slide_title_fn=lambda _slide: {"title": "识别标题", "shape_id": 5},
        extract_title_shape_extra_lines_fn=lambda _slide, _info: ["副标题", "要点"],
    )

    assert ctx["title_text"] == "识别标题"
    assert ctx["skip_map"] == {5: "识别标题"}
    assert ctx["exclude_ids"] == {5}
    assert ctx["extra_lines"] == ["副标题", "要点"]


def test_normalize_md_link_path_core():
    assert core.normalize_md_link_path_core(r"img\foo\bar.png") == "img/foo/bar.png"


def test_build_image_placeholder_markdown_core():
    assert core.build_image_placeholder_markdown_core(alt_text="方括号]测试") == "![图片: 方括号\\]测试]"
    assert core.build_image_placeholder_markdown_core(alt_text="") == "![图片]"


def test_next_export_image_path_core():
    image_ctx = {"counter": 0, "dir": r"D:\tmp\img"}

    class _S:
        Id = 12

    p = core.next_export_image_path_core(
        image_ctx,
        image_loc="S1/R1/SH12",
        shape=_S(),
        safe_shape_id_fn=lambda s: int(s.Id),
    )

    assert image_ctx["counter"] == 1
    assert str(p).endswith("img_0001_S1_R1_SH12_s12.png")


def test_build_image_extract_context_core():
    workdir = _make_workspace_dir()
    output_path = workdir / "result.md"

    enabled_ctx = core.build_image_extract_context_core(str(output_path), extract_images=True, image_dir="img")
    disabled_ctx = core.build_image_extract_context_core(str(output_path), extract_images=False, image_dir="img")

    assert enabled_ctx["enabled"] is True
    assert Path(enabled_ctx["dir"]).name == "img"
    assert Path(enabled_ctx["md_dir"]).resolve() == workdir.resolve()
    assert disabled_ctx == {"enabled": False, "dir": None, "md_dir": str(workdir.resolve()), "counter": 0}


def test_export_shape_image_markdown_core_success():
    workdir = _make_workspace_dir()
    image_ctx = {"enabled": True, "dir": str(workdir / "img"), "md_dir": str(workdir), "counter": 0}

    class _ImageShape:
        Id = 7

        def __init__(self):
            self.exports = []

        def Export(self, path, fmt):
            self.exports.append((path, fmt))
            Path(path).write_bytes(b"png")

    shape = _ImageShape()

    md = core.export_shape_image_markdown_core(
        shape,
        image_ctx=image_ctx,
        image_loc="S1/R1/SH7",
        read_shape_alt_text_fn=lambda _s: "封面图",
        build_image_placeholder_markdown_fn=lambda **kwargs: core.build_image_placeholder_markdown_core(**kwargs),
        next_export_image_path_fn=lambda ctx, image_loc=None, shape=None: core.next_export_image_path_core(
            ctx, image_loc=image_loc, shape=shape, safe_shape_id_fn=lambda s: int(s.Id)
        ),
        wait_com_fn=lambda action, _timeout, _context: action(),
    )

    assert md.startswith("![图片: 封面图](img/")
    assert md.endswith(".png)")
    assert image_ctx["counter"] == 1
    assert shape.exports and shape.exports[0][1] == 2


def test_export_shape_image_markdown_core_disabled_returns_placeholder():
    image_ctx = {"enabled": False, "dir": None, "md_dir": "", "counter": 0}

    md = core.export_shape_image_markdown_core(
        shape=SimpleNamespace(),
        image_ctx=image_ctx,
        image_loc="S1/R1/SH7",
        read_shape_alt_text_fn=lambda _s: "封面图",
        build_image_placeholder_markdown_fn=lambda **kwargs: core.build_image_placeholder_markdown_core(**kwargs),
        next_export_image_path_fn=lambda *_args, **_kwargs: (_ for _ in ()).throw(RuntimeError("不应调用")),
    )

    assert md == "![图片: 封面图]"


def test_process_shape_rows_core():
    normal = _Shape(shape_type=1)
    embedded_ppt = _Shape(shape_type=7, prog_id="PowerPoint.Show.12")
    embedded_excel = _Shape(shape_type=7, prog_id="Excel.Sheet.12")

    calls = []

    def _row_renderer(row_shapes, skip_first_para_by_shape_id=None, image_ctx=None, loc_prefix=None):
        calls.append(
            {
                "count": len(row_shapes),
                "skip_map": dict(skip_first_para_by_shape_id or {}),
                "image_ctx": dict(image_ctx or {}),
                "loc_prefix": loc_prefix,
            }
        )
        return ["正文行"]

    lines, embedded_shapes = core.process_shape_rows_core(
        [[normal, embedded_ppt, embedded_excel]],
        slide_loc="S2",
        row_renderer_fn=_row_renderer,
        skip_map={1: "标题"},
        image_ctx={"enabled": True},
        embedded_object_line_fn=lambda pid: f"embedded-object: {pid}",
    )

    assert lines == ["embedded-object: Excel.Sheet.12", "正文行"]
    assert embedded_shapes == [embedded_ppt]
    assert calls == [
        {"count": 1, "skip_map": {1: "标题"}, "image_ctx": {"enabled": True}, "loc_prefix": "S2/R1"},
    ]


def test_render_row_text_lines_builds_row_loc_prefix():
    calls = {}

    def _fake_render_shape_row_fn(row_shapes, skip_first_para_by_shape_id=None, image_ctx=None, loc_prefix=None):
        calls["row_shapes"] = list(row_shapes)
        calls["skip_map"] = dict(skip_first_para_by_shape_id or {})
        calls["image_ctx"] = dict(image_ctx or {})
        calls["loc_prefix"] = loc_prefix
        return ["line-1", "line-2"]

    lines = core.render_row_text_lines(
        row_shapes=["a", "b"],
        row_idx=2,
        slide_loc="S3",
        render_shape_row_fn=_fake_render_shape_row_fn,
        skip_map={10: "标题"},
        image_ctx={"enabled": True},
    )

    assert lines == ["line-1", "line-2"]
    assert calls["row_shapes"] == ["a", "b"]
    assert calls["skip_map"] == {10: "标题"}
    assert calls["image_ctx"] == {"enabled": True}
    assert calls["loc_prefix"] == "S3/R2"


def test_render_shape_row_with_number_merge():
    class _TextShape:
        def __init__(self, sid, text):
            self._sid = sid
            self.text = text

    class _OtherShape:
        def __init__(self, sid, text):
            self._sid = sid
            self.text = text

    number_shape = _TextShape(1, "2")
    title_shape = _TextShape(2, "流水线优势")
    other_shape = _OtherShape(3, "补充说明")

    calls = []

    def _safe_shape_id(shape):
        return getattr(shape, "_sid", None)

    def _single_line_text(shape):
        if isinstance(shape, _TextShape):
            return shape.text
        return None

    def _escape(text):
        return str(text)

    def _extract_text(shape, skip_first_para_text=None, image_ctx=None, image_loc=None):
        calls.append(
            {
                "sid": getattr(shape, "_sid", None),
                "skip_first_para_text": skip_first_para_text,
                "image_ctx": image_ctx,
                "image_loc": image_loc,
            }
        )
        return [shape.text]

    lines = core.render_shape_row_with_number_merge(
        [number_shape, title_shape, other_shape],
        skip_first_para_by_shape_id=None,
        image_ctx={"enabled": True},
        loc_prefix="S1/R1",
        safe_shape_id_fn=_safe_shape_id,
        get_single_line_plain_text_fn=_single_line_text,
        escape_md_text_line_fn=_escape,
        extract_text_from_shape_fn=_extract_text,
    )

    assert lines == ["2. 流水线优势", "补充说明"]
    assert calls == [
        {"sid": 3, "skip_first_para_text": None, "image_ctx": {"enabled": True}, "image_loc": "S1/R1/SH3"},
    ]


def test_extract_text_from_shape_core_image_branch():
    image_shape = SimpleNamespace(Type=13, HasTextFrame=False, HasTable=False)
    calls = {}

    def _export(shape, image_ctx=None, image_loc=None):
        calls["shape"] = shape
        calls["image_ctx"] = image_ctx
        calls["image_loc"] = image_loc
        return "![图片](img/x.png)"

    lines = core.extract_text_from_shape_core(
        image_shape,
        image_ctx={"enabled": True},
        image_loc="S1/R1/SH1",
        export_shape_image_markdown_fn=_export,
    )

    assert lines == ["![图片](img/x.png)"]
    assert calls["shape"] is image_shape
    assert calls["image_ctx"] == {"enabled": True}
    assert calls["image_loc"] == "S1/R1/SH1"
