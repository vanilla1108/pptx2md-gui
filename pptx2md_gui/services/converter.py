"""后台处理的转换工作线程。"""

import logging
import queue
import threading
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, List, Optional, Tuple

from pptx2md import convert
from pptx2md.types import ConversionConfig

from .config_bridge import build_config


@dataclass
class ConversionResults:
    """批量转换操作的结果。"""
    success_count: int = 0
    failed_count: int = 0
    total_count: int = 0
    failed_files: List[Tuple[Path, str]] = field(default_factory=list)


class QueueLogHandler(logging.Handler):
    """将日志记录放入队列的日志处理器。"""

    def __init__(self, log_queue: queue.Queue):
        super().__init__()
        self.log_queue = log_queue

    def emit(self, record):
        try:
            level = record.levelname
            if level == "INFO" and "完成" in record.getMessage():
                level = "SUCCESS"
            self.log_queue.put((level, record.getMessage()))
        except Exception:
            self.handleError(record)


class ConversionWorker(threading.Thread):
    """文件转换的后台工作线程。"""

    def __init__(
        self,
        files: List[Path],
        params: dict,
        log_queue: queue.Queue,
        progress_callback: Callable[[float, str], None],
        cancel_event: threading.Event,
        on_complete: Optional[Callable[[ConversionResults], None]] = None,
    ):
        """初始化转换工作线程。

        参数:
            files: 待转换的 PPTX 文件列表。
            params: 来自 ParamsPanel.get_params() 的 GUI 参数字典。
            log_queue: 日志消息队列（级别, 消息）。
            progress_callback: 进度更新回调（值 0-1, 状态）。
            cancel_event: 取消信号事件。
            on_complete: 转换完成时的可选回调。
        """
        super().__init__(daemon=True)
        self.files = files
        self.params = params
        self.log_queue = log_queue
        self.progress_callback = progress_callback
        self.cancel_event = cancel_event
        self.on_complete = on_complete
        self._results = ConversionResults(total_count=len(files))

    def run(self):
        """执行转换任务。"""
        self.log_queue.put(("INFO", f"开始转换任务，共 {len(self.files)} 个文件"))

        # 批次级环境检查：若有 .ppt 文件则预检一次
        ppt_env_ok, ppt_env_reason = True, ""
        ppt_files = [f for f in self.files if f.suffix.lower() == ".ppt"]
        if ppt_files:
            from pptx2md.ppt_legacy import check_environment
            ppt_env_ok, ppt_env_reason = check_environment(strict=True)
            if not ppt_env_ok:
                for f in ppt_files:
                    self.log_queue.put(("ERROR", f"[PPT] {f.name}: {ppt_env_reason}"))
                    self._results.failed_count += 1
                    self._results.failed_files.append((f, ppt_env_reason))

        for file_idx, file_path in enumerate(self.files):
            if self.cancel_event.is_set():
                self.log_queue.put(("WARNING", "用户取消转换"))
                break

            ext = file_path.suffix.lower()

            # 跳过环境检查失败的 .ppt
            if ext == ".ppt" and not ppt_env_ok:
                overall_progress = (file_idx + 1) / len(self.files)
                self.progress_callback(overall_progress, f"已完成 {file_idx + 1}/{len(self.files)}")
                continue

            self.log_queue.put(("INFO", f"正在处理: {file_path.name}"))

            try:
                if ext == ".ppt":
                    self._convert_ppt(file_path, file_idx)
                else:
                    self._convert_pptx(file_path, file_idx)
            except Exception as e:
                self._results.failed_count += 1
                self._results.failed_files.append((file_path, str(e)))
                self.log_queue.put(("ERROR", f"{file_path.name} 转换失败: {e}"))

            # 更新总体进度
            overall_progress = (file_idx + 1) / len(self.files)
            self.progress_callback(overall_progress, f"已完成 {file_idx + 1}/{len(self.files)}")

        # 最终汇总
        if not self.cancel_event.is_set():
            self.log_queue.put(("INFO", "═" * 40))
            self.log_queue.put((
                "INFO",
                f"转换完成！成功: {self._results.success_count}, "
                f"失败: {self._results.failed_count}, 总计: {self._results.total_count}"
            ))

        if self.on_complete:
            self.on_complete(self._results)

    def _convert_pptx(self, file_path: Path, file_idx: int):
        """走 pptx2md 原管道。"""
        config = build_config(file_path, self.params)

        def slide_progress(current, total, name):
            file_progress = file_idx / len(self.files)
            slide_progress_val = current / total / len(self.files)
            overall = file_progress + slide_progress_val
            status = f"正在处理: {file_path.name} ({current}/{total})"
            self.progress_callback(overall, status)

        convert(
            config,
            progress_callback=slide_progress,
            cancel_event=self.cancel_event,
            disable_tqdm=True,
        )

        if not self.cancel_event.is_set():
            self._results.success_count += 1
            output_name = config.output_path.name
            self.log_queue.put(("SUCCESS", f"{file_path.name} 转换完成 → {output_name}"))

    def _convert_ppt(self, file_path: Path, file_idx: int):
        """走 ppt_legacy COM 管道。"""
        from pptx2md.ppt_legacy import convert_ppt
        from .config_bridge import build_ppt_config

        # 格式警告
        fmt = self.params.get("output_format", "markdown").lower()
        if fmt != "markdown":
            self.log_queue.put((
                "WARNING",
                f"[PPT] {file_path.name}: PPT 转换仅支持 Markdown 输出，已自动切换"
            ))

        config = build_ppt_config(file_path, self.params)

        def log_cb(level, msg):
            self.log_queue.put((level, msg))

        def progress_cb(current, total, name):
            file_progress = file_idx / len(self.files)
            slide_progress_val = current / total / len(self.files)
            overall = file_progress + slide_progress_val
            status = f"正在处理: {file_path.name} ({current}/{total})"
            self.progress_callback(overall, status)

        success = convert_ppt(
            config,
            log_callback=log_cb,
            progress_callback=progress_cb,
            cancel_event=self.cancel_event,
        )

        if self.cancel_event.is_set():
            return  # 取消不计为成功或失败

        if success:
            self._results.success_count += 1
            output_name = Path(config.output_path).name if config.output_path else file_path.stem + ".md"
            self.log_queue.put(("SUCCESS", f"[PPT] {file_path.name} 转换完成 → {output_name}"))
        else:
            self._results.failed_count += 1
            self._results.failed_files.append((file_path, "转换返回失败"))
            self.log_queue.put(("ERROR", f"[PPT] {file_path.name}: 转换失败"))

    def get_results(self) -> ConversionResults:
        """获取转换结果。"""
        return self._results
