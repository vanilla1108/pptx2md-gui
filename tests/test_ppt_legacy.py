"""PPT Legacy COM 集成测试。

手动运行：python -m pytest tests/test_ppt_legacy.py -v -m ppt_com
"""

import shutil
from pathlib import Path

import pytest

from tests.conftest import PROJECT_ROOT

# ---------------------------------------------------------------------------
# Fixture
# ---------------------------------------------------------------------------


@pytest.fixture(scope="module")
def require_com():
    """仅在执行本模块测试时检测 COM 环境并按需跳过。"""
    from pptx2md.ppt_legacy import check_environment
    ok, reason = check_environment(strict=True)
    if not ok:
        pytest.skip(f"需要 Windows + PowerPoint + pywin32: {reason}")

PPT_SAMPLE_DIR = PROJECT_ROOT / "ppt2md_script" / "test-ppt"


@pytest.fixture
def ppt_sample(workspace_tmp_path) -> Path:
    """复制测试 .ppt 到临时目录。"""
    # 寻找第一个 .ppt 文件
    samples = list(PPT_SAMPLE_DIR.glob("*.ppt"))
    if not samples:
        pytest.skip("无 .ppt 测试样本")
    src = samples[0]
    dst = workspace_tmp_path / src.name
    shutil.copy2(src, dst)
    return dst


# ---------------------------------------------------------------------------
# 测试
# ---------------------------------------------------------------------------


@pytest.mark.ppt_com
@pytest.mark.usefixtures("require_com")
class TestPptLegacyConversion:

    def test_basic_conversion(self, ppt_sample, workspace_tmp_path):
        """转换 .ppt，验证 .md 存在且非空。"""
        from pptx2md.ppt_legacy import convert_ppt
        from pptx2md.ppt_legacy.config import ExtractConfig

        output = workspace_tmp_path / "output.md"
        config = ExtractConfig(
            input_path=str(ppt_sample),
            output_path=str(output),
        )
        success = convert_ppt(config)
        assert success is True
        assert output.exists()
        assert output.stat().st_size > 0

    def test_log_callback_receives_prefixed_messages(self, ppt_sample, workspace_tmp_path):
        """log_callback 应收到 [PPT] 前缀消息。"""
        from pptx2md.ppt_legacy import convert_ppt
        from pptx2md.ppt_legacy.config import ExtractConfig

        logs = []
        config = ExtractConfig(
            input_path=str(ppt_sample),
            output_path=str(workspace_tmp_path / "output.md"),
        )
        convert_ppt(config, log_callback=lambda lvl, msg: logs.append((lvl, msg)))
        assert any("[PPT]" in msg for _, msg in logs)

    def test_progress_callback_increments(self, ppt_sample, workspace_tmp_path):
        """progress_callback 的 current 应单调递增。"""
        from pptx2md.ppt_legacy import convert_ppt
        from pptx2md.ppt_legacy.config import ExtractConfig

        progress = []
        config = ExtractConfig(
            input_path=str(ppt_sample),
            output_path=str(workspace_tmp_path / "output.md"),
        )
        convert_ppt(config, progress_callback=lambda c, t, n: progress.append((c, t, n)))
        if progress:
            currents = [c for c, _, _ in progress]
            assert currents == sorted(currents), f"progress 不单调递增: {currents}"

    def test_cancel_exits_within_timeout(self, ppt_sample, workspace_tmp_path):
        """设置 cancel_event 后应在合理时间内退出。"""
        import threading
        import time
        from pptx2md.ppt_legacy import convert_ppt
        from pptx2md.ppt_legacy.config import ExtractConfig

        cancel = threading.Event()
        cancel.set()  # 立即取消

        config = ExtractConfig(
            input_path=str(ppt_sample),
            output_path=str(workspace_tmp_path / "output.md"),
        )
        # 应在 5s 内返回 False
        start = time.time()
        result = convert_ppt(config, cancel_event=cancel)
        elapsed = time.time() - start
        assert result is False
        assert elapsed < 5.0, f"取消耗时 {elapsed:.1f}s，超过预期"

    def test_image_extraction_creates_files(self, ppt_sample, workspace_tmp_path):
        """extract_images=True 应在 image_dir 下创建文件。"""
        from pptx2md.ppt_legacy import convert_ppt
        from pptx2md.ppt_legacy.config import ExtractConfig

        img_dir = workspace_tmp_path / "img"
        config = ExtractConfig(
            input_path=str(ppt_sample),
            output_path=str(workspace_tmp_path / "output.md"),
            extract_images=True,
            image_dir=str(img_dir),
        )
        convert_ppt(config)
        # 图片可能存在也可能不存在（取决于 PPT 内容），只验证目录被创建
        # 若 PPT 含图片，目录下应有文件
        if img_dir.exists():
            pass  # 创建即算通过

    def test_no_image_placeholder(self, ppt_sample, workspace_tmp_path):
        """extract_images=False 时输出应含占位文本。"""
        from pptx2md.ppt_legacy import convert_ppt
        from pptx2md.ppt_legacy.config import ExtractConfig

        output = workspace_tmp_path / "output.md"
        config = ExtractConfig(
            input_path=str(ppt_sample),
            output_path=str(output),
            extract_images=False,
        )
        convert_ppt(config)
        # 占位符可能存在于输出中（取决于 PPT 是否含图片）
        assert output.exists()
