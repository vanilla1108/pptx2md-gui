"""后台处理的转换工作线程（父线程调度 + 子进程执行）。"""

import logging
import multiprocessing as mp
import queue
import threading
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable, List, Optional, Tuple

from .config_bridge import build_config, build_ppt_config

_SUBPROCESS_LOG = "log"
_SUBPROCESS_PROGRESS = "slide_progress"
_SUBPROCESS_RESULT = "result"


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


class _SubprocessQueueLogHandler(logging.Handler):
    """子进程日志桥接：把 logging 输出转发回父进程。"""

    def __init__(self, msg_queue: Any):
        super().__init__()
        self._msg_queue = msg_queue

    def emit(self, record):
        try:
            message = record.getMessage()
            self._msg_queue.put((_SUBPROCESS_LOG, record.levelname, message))
        except Exception:
            # 子进程内日志转发失败时不影响主流程
            pass


def _setup_subprocess_logging(msg_queue: Any):
    """重置子进程日志，将根日志输出定向到父进程消息队列。"""

    root_logger = logging.getLogger()
    for handler in root_logger.handlers[:]:
        root_logger.removeHandler(handler)

    root_logger.setLevel(logging.INFO)
    root_logger.addHandler(_SubprocessQueueLogHandler(msg_queue))


def _convert_single_pptx_file(file_path: Path, params: dict, msg_queue: Any):
    from pptx2md import convert

    config = build_config(file_path, params)

    def progress_cb(current, total, name):
        msg_queue.put((_SUBPROCESS_PROGRESS, current, total, name))

    convert(
        config,
        progress_callback=progress_cb,
        cancel_event=None,
        disable_tqdm=True,
    )

    output_name = config.output_path.name
    msg_queue.put((_SUBPROCESS_LOG, "SUCCESS", f"{file_path.name} 转换完成 → {output_name}"))
    msg_queue.put((_SUBPROCESS_RESULT, {"success": True, "output_name": output_name, "error": ""}))


def _convert_single_ppt_file(file_path: Path, params: dict, msg_queue: Any):
    from pptx2md.ppt_legacy import check_environment, convert_ppt

    env_ok, env_reason = check_environment(strict=True)
    if not env_ok:
        msg_queue.put((_SUBPROCESS_LOG, "ERROR", f"[PPT] {file_path.name}: {env_reason}"))
        msg_queue.put((_SUBPROCESS_RESULT, {"success": False, "output_name": "", "error": env_reason}))
        return

    fmt = str(params.get("output_format", "markdown")).lower()
    if fmt != "markdown":
        msg_queue.put((
            _SUBPROCESS_LOG,
            "WARNING",
            f"[PPT] {file_path.name}: PPT 转换仅支持 Markdown 输出，已自动切换",
        ))

    config = build_ppt_config(file_path, params)

    def log_cb(level, message):
        msg_queue.put((_SUBPROCESS_LOG, level, message))

    def progress_cb(current, total, name):
        msg_queue.put((_SUBPROCESS_PROGRESS, current, total, name))

    success = convert_ppt(
        config,
        log_callback=log_cb,
        progress_callback=progress_cb,
        cancel_event=None,
    )

    if success:
        output_name = Path(config.output_path).name if config.output_path else file_path.stem + ".md"
        msg_queue.put((_SUBPROCESS_LOG, "SUCCESS", f"[PPT] {file_path.name} 转换完成 → {output_name}"))
        msg_queue.put((_SUBPROCESS_RESULT, {"success": True, "output_name": output_name, "error": ""}))
        return

    msg_queue.put((_SUBPROCESS_LOG, "ERROR", f"[PPT] {file_path.name}: 转换失败"))
    msg_queue.put((_SUBPROCESS_RESULT, {"success": False, "output_name": "", "error": "转换返回失败"}))


def _convert_single_file_subprocess(file_path_str: str, params: dict, msg_queue: Any):
    """子进程入口：执行单文件转换并把消息回传父进程。"""

    file_path = Path(file_path_str)
    _setup_subprocess_logging(msg_queue)

    try:
        if file_path.suffix.lower() == ".ppt":
            _convert_single_ppt_file(file_path, params, msg_queue)
        else:
            _convert_single_pptx_file(file_path, params, msg_queue)
    except Exception as exc:
        error_msg = str(exc) or exc.__class__.__name__
        msg_queue.put((_SUBPROCESS_LOG, "ERROR", f"{file_path.name} 转换失败: {error_msg}"))
        msg_queue.put((_SUBPROCESS_RESULT, {"success": False, "output_name": "", "error": error_msg}))


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
        """初始化转换工作线程。"""

        super().__init__(daemon=True)
        self.files = files
        self.params = params
        self.log_queue = log_queue
        self.progress_callback = progress_callback
        self.cancel_event = cancel_event
        self.on_complete = on_complete
        self._results = ConversionResults(total_count=len(files))
        self._mp_ctx = mp.get_context("spawn")

    def run(self):
        """执行转换任务。"""

        self.log_queue.put(("INFO", f"开始转换任务，共 {len(self.files)} 个文件"))

        for file_idx, file_path in enumerate(self.files):
            if self.cancel_event.is_set():
                self.log_queue.put(("WARNING", "用户取消转换"))
                break

            self.log_queue.put(("INFO", f"正在处理: {file_path.name}"))

            try:
                success, detail = self._run_single_file_in_subprocess(file_path, file_idx)
            except Exception as exc:
                self.log_queue.put((
                    "WARNING",
                    f"子进程不可用，回退到线程内转换：{exc}",
                ))
                success, detail = self._run_single_file_in_process(file_path, file_idx)

            if success is None:
                self.log_queue.put(("WARNING", "用户取消转换"))
                break

            if success:
                self._results.success_count += 1
            else:
                self._results.failed_count += 1
                self._results.failed_files.append((file_path, detail))
                if detail.startswith("子进程"):
                    self.log_queue.put(("ERROR", f"{file_path.name} 转换失败: {detail}"))

            overall_progress = (file_idx + 1) / len(self.files)
            self.progress_callback(overall_progress, f"已完成 {file_idx + 1}/{len(self.files)}")

        if not self.cancel_event.is_set():
            self.log_queue.put(("INFO", "═" * 40))
            self.log_queue.put((
                "INFO",
                f"转换完成！成功: {self._results.success_count}, "
                f"失败: {self._results.failed_count}, 总计: {self._results.total_count}",
            ))

        if self.on_complete:
            self.on_complete(self._results)

    def _run_single_file_in_subprocess(self, file_path: Path, file_idx: int) -> Tuple[Optional[bool], str]:
        """在子进程中执行单文件转换。

        返回:
            (True, output_name) 表示成功；
            (False, error_message) 表示失败；
            (None, "") 表示用户取消。
        """

        try:
            msg_queue = self._mp_ctx.Queue()
            process = self._mp_ctx.Process(
                target=_convert_single_file_subprocess,
                args=(str(file_path), self.params, msg_queue),
                daemon=False,
            )
            process.start()
        except Exception as exc:
            raise RuntimeError(str(exc) or exc.__class__.__name__) from exc

        result_payload: Optional[dict] = None

        try:
            while process.is_alive():
                if self.cancel_event.is_set():
                    process.terminate()
                    process.join(timeout=5)
                    self._drain_subprocess_messages(msg_queue, file_path, file_idx)
                    return None, ""

                try:
                    message = msg_queue.get(timeout=0.1)
                except queue.Empty:
                    continue

                payload = self._forward_subprocess_message(message, file_path, file_idx)
                if payload is not None:
                    result_payload = payload

            process.join()

            drained_payload = self._drain_subprocess_messages(msg_queue, file_path, file_idx)
            if drained_payload is not None:
                result_payload = drained_payload
        finally:
            try:
                msg_queue.close()
                msg_queue.join_thread()
            except Exception:
                pass

        if result_payload and result_payload.get("success"):
            return True, str(result_payload.get("output_name", ""))

        error_message = ""
        if result_payload:
            error_message = str(result_payload.get("error") or "").strip()

        if not error_message:
            if process.exitcode is None:
                error_message = "子进程状态未知"
            elif process.exitcode != 0:
                error_message = f"子进程异常退出（exit code {process.exitcode}）"
            else:
                error_message = "子进程未返回结果"

        return False, error_message

    def _run_single_file_in_process(self, file_path: Path, file_idx: int) -> Tuple[Optional[bool], str]:
        """子进程不可用时的兜底路径：在线程内执行单文件转换。"""

        try:
            if file_path.suffix.lower() == ".ppt":
                return self._convert_single_ppt_in_process(file_path, file_idx)
            return self._convert_single_pptx_in_process(file_path, file_idx)
        except Exception as exc:
            error_message = str(exc) or exc.__class__.__name__
            self.log_queue.put(("ERROR", f"{file_path.name} 转换失败: {error_message}"))
            return False, error_message

    def _convert_single_pptx_in_process(self, file_path: Path, file_idx: int) -> Tuple[Optional[bool], str]:
        """线程内执行 .pptx 转换（兜底）。"""

        from pptx2md import convert

        config = build_config(file_path, self.params)

        def progress_cb(current, total, name):
            if total <= 0:
                return
            file_progress = file_idx / len(self.files)
            slide_progress_val = (current / total) / len(self.files)
            overall = file_progress + slide_progress_val
            status = f"正在处理: {file_path.name} ({int(current)}/{int(total)})"
            self.progress_callback(overall, status)

        convert(
            config,
            progress_callback=progress_cb,
            cancel_event=self.cancel_event,
            disable_tqdm=True,
        )

        if self.cancel_event.is_set():
            return None, ""

        output_name = config.output_path.name
        self.log_queue.put(("SUCCESS", f"{file_path.name} 转换完成 → {output_name}"))
        return True, output_name

    def _convert_single_ppt_in_process(self, file_path: Path, file_idx: int) -> Tuple[Optional[bool], str]:
        """线程内执行 .ppt 转换（兜底）。"""

        from pptx2md.ppt_legacy import check_environment, convert_ppt

        env_ok, env_reason = check_environment(strict=True)
        if not env_ok:
            self.log_queue.put(("ERROR", f"[PPT] {file_path.name}: {env_reason}"))
            return False, env_reason

        fmt = str(self.params.get("output_format", "markdown")).lower()
        if fmt != "markdown":
            self.log_queue.put((
                "WARNING",
                f"[PPT] {file_path.name}: PPT 转换仅支持 Markdown 输出，已自动切换",
            ))

        config = build_ppt_config(file_path, self.params)

        def log_cb(level, message):
            self.log_queue.put((level, message))

        def progress_cb(current, total, name):
            if total <= 0:
                return
            file_progress = file_idx / len(self.files)
            slide_progress_val = (current / total) / len(self.files)
            overall = file_progress + slide_progress_val
            status = f"正在处理: {file_path.name} ({int(current)}/{int(total)})"
            self.progress_callback(overall, status)

        success = convert_ppt(
            config,
            log_callback=log_cb,
            progress_callback=progress_cb,
            cancel_event=self.cancel_event,
        )

        if self.cancel_event.is_set():
            return None, ""

        if success:
            output_name = Path(config.output_path).name if config.output_path else file_path.stem + ".md"
            self.log_queue.put(("SUCCESS", f"[PPT] {file_path.name} 转换完成 → {output_name}"))
            return True, output_name

        self.log_queue.put(("ERROR", f"[PPT] {file_path.name}: 转换失败"))
        return False, "转换返回失败"

    def _drain_subprocess_messages(self, msg_queue: Any, file_path: Path, file_idx: int) -> Optional[dict]:
        """清空子进程消息队列，返回最后一个 result payload。"""

        result_payload: Optional[dict] = None
        while True:
            try:
                message = msg_queue.get_nowait()
            except queue.Empty:
                break

            payload = self._forward_subprocess_message(message, file_path, file_idx)
            if payload is not None:
                result_payload = payload

        return result_payload

    def _forward_subprocess_message(
        self,
        message: Any,
        file_path: Path,
        file_idx: int,
    ) -> Optional[dict]:
        """处理子进程回传消息并更新 GUI 队列/进度。"""

        if not isinstance(message, tuple) or not message:
            return None

        kind = message[0]

        if kind == _SUBPROCESS_LOG and len(message) >= 3:
            level = str(message[1])
            content = str(message[2])
            if level == "INFO" and "完成" in content:
                level = "SUCCESS"
            self.log_queue.put((level, content))
            return None

        if kind == _SUBPROCESS_PROGRESS and len(message) >= 4:
            try:
                current = float(message[1])
                total = float(message[2])
            except (TypeError, ValueError):
                return None

            if total <= 0:
                return None

            file_progress = file_idx / len(self.files)
            slide_progress_val = (current / total) / len(self.files)
            overall = file_progress + slide_progress_val
            status = f"正在处理: {file_path.name} ({int(current)}/{int(total)})"
            self.progress_callback(overall, status)
            return None

        if kind == _SUBPROCESS_RESULT and len(message) >= 2 and isinstance(message[1], dict):
            return message[1]

        return None

    def get_results(self) -> ConversionResults:
        """获取转换结果。"""

        return self._results
