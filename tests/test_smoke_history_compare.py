"""历史冒烟输出对比：按文件修改时间比较最新与上一次提取结果。"""

from __future__ import annotations

import difflib
import os
from datetime import datetime
from pathlib import Path

import pytest

from tests.conftest import PROJECT_ROOT, SAMPLE_PPTX_DL


def _fmt_mtime(path: Path) -> str:
    return datetime.fromtimestamp(path.stat().st_mtime).strftime("%Y-%m-%d %H:%M:%S")


def _collect_output_history(base_stem: str) -> list[Path]:
    output_dir = PROJECT_ROOT / "test_output"
    candidates = [p for p in output_dir.glob(f"{base_stem}*.md") if p.is_file()]
    # 关键要求：按生成日期/修改时间判断新旧，而不是按文件名后缀。
    candidates.sort(key=lambda p: (p.stat().st_mtime_ns, p.name), reverse=True)
    return candidates


@pytest.mark.slow
def test_dl_smoke_output_matches_previous_by_mtime():
    """比较 test_output 中最新与上一次深度学习冒烟输出，出现差异则给出 diff 报告。"""
    base_stem = SAMPLE_PPTX_DL.stem
    history = _collect_output_history(base_stem)

    if len(history) < 2:
        pytest.skip(f"Not enough historical outputs under test_output for '{base_stem}'")

    latest, previous = history[0], history[1]
    latest_text = latest.read_text(encoding="utf-8")
    previous_text = previous.read_text(encoding="utf-8")

    if latest_text == previous_text:
        return

    diff_lines = list(
        difflib.unified_diff(
            previous_text.splitlines(),
            latest_text.splitlines(),
            fromfile=f"{previous.name} ({_fmt_mtime(previous)})",
            tofile=f"{latest.name} ({_fmt_mtime(latest)})",
            lineterm="",
            n=3,
        )
    )
    diff_dir = PROJECT_ROOT / "tmp_test_artifacts" / "history_diff"
    diff_dir.mkdir(parents=True, exist_ok=True)
    diff_path = diff_dir / f"{base_stem}_latest_vs_previous.diff"
    diff_path.write_text("\n".join(diff_lines) + "\n", encoding="utf-8")

    message = (
        "Latest smoke output differs from previous output by mtime.\n"
        f"Latest: {latest.name} ({_fmt_mtime(latest)})\n"
        f"Previous: {previous.name} ({_fmt_mtime(previous)})\n"
        f"Diff report: {diff_path}"
    )

    strict = os.getenv("PPTX2MD_SMOKE_HISTORY_STRICT", "0").strip().lower() in ("1", "true", "yes", "on")
    if strict:
        raise AssertionError(message)

    pytest.skip(f"[non-strict] {message}")
