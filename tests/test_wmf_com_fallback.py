from pathlib import Path

from PIL import Image

import pptx2md.parser as parser
from pptx2md.types import ConversionConfig


class _FakeShape:
    left = 120
    top = 180
    width = 400
    height = 240


class _FakeComSession:
    def __init__(self):
        self.open_path = None
        self.export_path = None

    def ensure_open(self, pptx_path: str):
        self.open_path = pptx_path

    def export_slide_png(self, slide_idx: int, slide_png_path: str, width_px: int, height_px: int) -> bool:
        self.export_path = slide_png_path
        slide_path = Path(slide_png_path)
        slide_path.parent.mkdir(parents=True, exist_ok=True)
        Image.new("RGB", (width_px, height_px), color="white").save(slide_path)
        return True


def test_wmf_com_fallback_uses_absolute_paths(monkeypatch, tmp_path):
    monkeypatch.chdir(tmp_path)
    fake_session = _FakeComSession()
    monkeypatch.setattr(parser, "_PPT_COM_SESSION", fake_session)
    monkeypatch.setattr(parser, "_get_slide_size_emu", lambda _pptx_path: (1000, 1000))

    config = ConversionConfig(
        pptx_path=Path("relative_input.pptx"),
        output_path=Path("out.md"),
        image_dir=Path("img"),
    )
    output_path = Path("img") / "wmf_raster.png"

    ok = parser._convert_wmf_via_powerpoint_slide_export(config, _FakeShape(), 1, str(output_path))

    assert ok is True
    assert fake_session.open_path is not None and Path(fake_session.open_path).is_absolute()
    assert fake_session.export_path is not None and Path(fake_session.export_path).is_absolute()
    assert output_path.exists()
