"""DropZone 拖拽数据解析测试。"""

from pathlib import Path

from pptx2md_gui.components.drop_zone import DropZone


class _TkSplitOk:
    """模拟 tk.splitlist 正常返回。"""

    @staticmethod
    def splitlist(data: str):
        return ("C:/tmp/a.pptx", "C:/tmp/with space/b.ppt")


class _TkSplitFail:
    """模拟 tk.splitlist 异常，触发回退解析。"""

    @staticmethod
    def splitlist(data: str):
        raise RuntimeError("split failed")


def _build_drop_zone_for_parse(fake_tk):
    zone = object.__new__(DropZone)
    zone.tk = fake_tk
    return zone


def test_parse_drop_data_supports_multi_files_via_splitlist():
    zone = _build_drop_zone_for_parse(_TkSplitOk())
    files = zone._parse_drop_data("{C:/tmp/a.pptx} {C:/tmp/with space/b.ppt}")
    assert files == [Path("C:/tmp/a.pptx"), Path("C:/tmp/with space/b.ppt")]


def test_parse_drop_data_fallback_handles_braced_multi_files():
    zone = _build_drop_zone_for_parse(_TkSplitFail())
    files = zone._parse_drop_data("{C:/tmp/a.pptx} {C:/tmp/with space/b.ppt}")
    assert files == [Path("C:/tmp/a.pptx"), Path("C:/tmp/with space/b.ppt")]


def test_parse_drop_data_empty_input():
    zone = _build_drop_zone_for_parse(_TkSplitOk())
    assert zone._parse_drop_data("") == []
