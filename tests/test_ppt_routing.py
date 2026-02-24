"""PPT 路由与配置的单元测试。"""

from pathlib import Path
from unittest.mock import patch

import pytest


class TestExtractConfig:
    """ExtractConfig 数据模型测试。"""

    def test_create_with_defaults(self):
        from pptx2md.ppt_legacy.config import ExtractConfig
        cfg = ExtractConfig(input_path="/tmp/test.ppt")
        assert cfg.input_path == "/tmp/test.ppt"
        assert cfg.output_path is None
        assert cfg.debug is False
        assert cfg.ui is True
        assert cfg.extract_images is True
        assert cfg.image_dir is None
        assert cfg.table_header == "first-row"

    def test_rejects_empty_input_path(self):
        from pptx2md.ppt_legacy.config import ExtractConfig
        with pytest.raises(ValueError, match="input_path"):
            ExtractConfig(input_path="")

    def test_rejects_invalid_table_header(self):
        from pptx2md.ppt_legacy.config import ExtractConfig
        with pytest.raises(ValueError, match="table_header"):
            ExtractConfig(input_path="/tmp/test.ppt", table_header="invalid")

    def test_frozen(self):
        from pptx2md.ppt_legacy.config import ExtractConfig
        import dataclasses
        cfg = ExtractConfig(input_path="/tmp/test.ppt")
        with pytest.raises(dataclasses.FrozenInstanceError):
            cfg.debug = True


class TestConversionCancelled:
    """ConversionCancelled 异常测试。"""

    def test_is_exception(self):
        from pptx2md.ppt_legacy.config import ConversionCancelled
        assert issubclass(ConversionCancelled, Exception)

    def test_not_base_exception(self):
        from pptx2md.ppt_legacy.config import ConversionCancelled
        # 确保不是 BaseException 子类（不绕过 cleanup）
        assert not issubclass(ConversionCancelled, KeyboardInterrupt)


class TestCheckEnvironment:
    """check_environment() 环境检测测试。"""

    @patch("platform.system", return_value="Linux")
    def test_non_windows(self, mock_sys):
        from pptx2md.ppt_legacy import check_environment
        ok, reason = check_environment()
        assert ok is False
        assert "Windows" in reason

    @patch("platform.system", return_value="Windows")
    def test_no_win32com(self, mock_sys):
        from pptx2md.ppt_legacy import check_environment
        with patch.dict("sys.modules", {"win32com": None, "win32com.client": None}):
            # 强制重新导入使 mock 生效
            ok, reason = check_environment()
            assert ok is False
            assert "pywin32" in reason


class TestFileValidators:
    """扩展名验证纯函数测试。"""

    def test_pptx_supported(self):
        from pptx2md_gui.utils.validators import is_supported_file
        assert is_supported_file(Path("test.pptx")) is True

    def test_ppt_supported(self):
        from pptx2md_gui.utils.validators import is_supported_file
        assert is_supported_file(Path("test.ppt")) is True

    def test_ppt_case_insensitive(self):
        from pptx2md_gui.utils.validators import is_supported_file
        assert is_supported_file(Path("test.PPT")) is True

    def test_odp_not_supported(self):
        from pptx2md_gui.utils.validators import is_supported_file
        assert is_supported_file(Path("test.odp")) is False

    def test_pdf_not_supported(self):
        from pptx2md_gui.utils.validators import is_supported_file
        assert is_supported_file(Path("test.pdf")) is False

    def test_has_ppt_in_mixed_list(self):
        from pptx2md_gui.utils.validators import has_ppt_in_list
        files = [Path("a.pptx"), Path("b.ppt"), Path("c.pptx")]
        assert has_ppt_in_list(files) is True

    def test_has_ppt_in_pptx_only_list(self):
        from pptx2md_gui.utils.validators import has_ppt_in_list
        files = [Path("a.pptx"), Path("b.pptx")]
        assert has_ppt_in_list(files) is False

    def test_has_ppt_empty_list(self):
        from pptx2md_gui.utils.validators import has_ppt_in_list
        assert has_ppt_in_list([]) is False


class TestCliRouting:
    """CLI 路由逻辑测试。"""

    def test_ppt_rejects_wiki(self):
        """传入 .ppt + --wiki 应报错。"""
        from pptx2md.__main__ import _check_ppt_format_conflict
        with pytest.raises(SystemExit):
            _check_ppt_format_conflict(wiki=True, mdk=False, qmd=False)

    def test_ppt_rejects_mdk(self):
        from pptx2md.__main__ import _check_ppt_format_conflict
        with pytest.raises(SystemExit):
            _check_ppt_format_conflict(wiki=False, mdk=True, qmd=False)

    def test_ppt_rejects_qmd(self):
        from pptx2md.__main__ import _check_ppt_format_conflict
        with pytest.raises(SystemExit):
            _check_ppt_format_conflict(wiki=False, mdk=False, qmd=True)

    def test_ppt_allows_markdown(self):
        """纯 Markdown 不应报错。"""
        from pptx2md.__main__ import _check_ppt_format_conflict
        # 不抛异常即通过
        _check_ppt_format_conflict(wiki=False, mdk=False, qmd=False)

    def test_build_ppt_config_basic(self):
        """_build_ppt_config 构建基本 ExtractConfig。"""
        from pptx2md.__main__ import _build_ppt_config
        from argparse import Namespace
        args = Namespace(
            pptx_path=Path("/tmp/test.ppt"),
            output=Path("/tmp/out.md"),
            image_dir=None,
            disable_image=False,
            ppt_debug=False,
            ppt_no_ui=False,
            ppt_table_header="first-row",
        )
        cfg = _build_ppt_config(args)
        assert cfg.input_path == str(Path("/tmp/test.ppt").resolve())
        assert cfg.extract_images is True
        assert cfg.debug is False

    def test_build_ppt_config_disable_image_reversed(self):
        """--disable-image 应映射为 extract_images=False。"""
        from pptx2md.__main__ import _build_ppt_config
        from argparse import Namespace
        args = Namespace(
            pptx_path=Path("/tmp/test.ppt"),
            output=Path("/tmp/out.md"),
            image_dir=None,
            disable_image=True,
            ppt_debug=False,
            ppt_no_ui=False,
            ppt_table_header="first-row",
        )
        cfg = _build_ppt_config(args)
        assert cfg.extract_images is False

    def test_build_ppt_config_default_output_in_cwd(self, tmp_path, monkeypatch):
        """未指定 --output 时，.ppt 默认输出到当前目录。"""
        from pptx2md.__main__ import _build_ppt_config
        from argparse import Namespace

        monkeypatch.chdir(tmp_path)
        args = Namespace(
            pptx_path=Path("demo.ppt"),
            output=None,
            image_dir=None,
            disable_image=False,
            ppt_debug=False,
            ppt_no_ui=False,
            ppt_table_header="first-row",
        )
        cfg = _build_ppt_config(args)
        assert cfg.output_path == str(tmp_path / "demo.md")

    def test_build_ppt_config_default_output_avoids_conflict(self, tmp_path, monkeypatch):
        """未指定 --output 时，若当前目录已有同名文件应自动避让。"""
        from pptx2md.__main__ import _build_ppt_config
        from argparse import Namespace

        monkeypatch.chdir(tmp_path)
        (tmp_path / "demo.md").write_text("existing", encoding="utf-8")
        args = Namespace(
            pptx_path=Path("demo.ppt"),
            output=None,
            image_dir=None,
            disable_image=False,
            ppt_debug=False,
            ppt_no_ui=False,
            ppt_table_header="first-row",
        )
        cfg = _build_ppt_config(args)
        assert cfg.output_path == str(tmp_path / "demo_1.md")

    def test_main_rejects_unsupported_extension(self, capsys, monkeypatch):
        """除 .ppt/.pptx 外应直接报错退出。"""
        from argparse import Namespace
        import pptx2md.__main__ as main_mod

        monkeypatch.setattr(main_mod, "parse_args", lambda: Namespace(pptx_path=Path("a.pdf")))
        with pytest.raises(SystemExit) as ex:
            main_mod.main()
        assert ex.value.code == 2
        assert "不支持的文件格式" in capsys.readouterr().err

    def test_warn_unsupported_ppt_params(self, capsys):
        """传入 .ppt 不支持的参数应输出警告。"""
        from pptx2md.__main__ import _warn_unsupported_ppt_params
        from argparse import Namespace
        args = Namespace(
            page=5,
            try_multi_column=True,
            title=None,
            image_width=None,
            min_block_size=15,
            keep_similar_titles=False,
            enable_color=True,
            disable_color=False,
            enable_escaping=True,
            disable_escaping=False,
            enable_notes=True,
            disable_notes=False,
            enable_slides=False,
            enable_slide_number=True,
            disable_slide_number=False,
            disable_wmf=False,
            compress_blank_lines=True,
        )
        _warn_unsupported_ppt_params(args)
        captured = capsys.readouterr()
        assert "--page" in captured.err
        assert "--try-multi-column" in captured.err
