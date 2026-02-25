"""后台处理的转换工作线程（父线程调度 + 子进程执行）。"""

import logging
import multiprocessing as mp
import queue
import threading
import time
from collections import deque
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Tuple

from .config_bridge import build_config, build_ppt_config

_SUBPROCESS_LOG = "log"
_SUBPROCESS_PROGRESS = "slide_progress"
_SUBPROCESS_RESULT = "result"
_MAX_AUTO_WORKERS = 4
_POLL_INTERVAL_SEC = 0.05


@dataclass
class ConversionResults:
    """批量转换操作的结果。"""

    success_count: int = 0
    failed_count: int = 0
    total_count: int = 0
    failed_files: List[Tuple[Path, str]] = field(default_factory=list)


@dataclass
class _RunningSubprocessTask:
    """父线程维护的子进程任务状态。"""

    file_idx: int
    file_path: Path
    process: Any
    msg_queue: Any
    result_payload: Optional[dict] = None


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
        self._file_progress: Dict[int, float] = {idx: 0.0 for idx in range(len(files))}
        self._max_workers = self._resolve_max_workers()

    def run(self):
        """执行转换任务。"""

        total = len(self.files)
        if total == 0:
            if self.on_complete:
                self.on_complete(self._results)
            return

        self.log_queue.put(("INFO", f"开始转换任务，共 {total} 个文件"))

        all_files = self._indexed_files()
        pptx_files, ppt_files = self._split_files_by_format(all_files)
        cancelled = self._run_conversion_plan(all_files, pptx_files, ppt_files)

        if cancelled or self.cancel_event.is_set():
            self.log_queue.put(("WARNING", "用户取消转换"))
        else:
            self.log_queue.put(("INFO", "═" * 40))
            self.log_queue.put((
                "INFO",
                f"转换完成！成功: {self._results.success_count}, "
                f"失败: {self._results.failed_count}, 总计: {self._results.total_count}",
            ))

        if self.on_complete:
            self.on_complete(self._results)

    def _indexed_files(self) -> List[Tuple[int, Path]]:
        """返回带原始索引的文件列表。"""
        return list(enumerate(self.files))

    @staticmethod
    def _split_files_by_format(
        indexed_files: List[Tuple[int, Path]],
    ) -> Tuple[List[Tuple[int, Path]], List[Tuple[int, Path]]]:
        """拆分为 .pptx 列表与 .ppt 列表（保持原始相对顺序）。"""

        pptx_files: List[Tuple[int, Path]] = []
        ppt_files: List[Tuple[int, Path]] = []
        for file_idx, file_path in indexed_files:
            if file_path.suffix.lower() == ".ppt":
                ppt_files.append((file_idx, file_path))
            else:
                pptx_files.append((file_idx, file_path))
        return pptx_files, ppt_files

    def _run_conversion_plan(
        self,
        all_files: List[Tuple[int, Path]],
        pptx_files: List[Tuple[int, Path]],
        ppt_files: List[Tuple[int, Path]],
    ) -> bool:
        """根据文件格式和并发配置选择执行策略，返回是否被取消。"""

        has_mixed = bool(pptx_files and ppt_files)
        use_concurrent = self._max_workers > 1
        target_pptx = pptx_files if has_mixed else all_files

        # 混合格式：先 pptx 再串行 ppt（避免 COM 冲突）
        if has_mixed:
            self.log_queue.put((
                "INFO",
                "检测到混合格式：将先并发转换 .pptx，再串行转换 .ppt（避免 COM 冲突）",
            ))

        # 仅 ppt 文件时强制串行
        if use_concurrent and ppt_files and not pptx_files:
            self.log_queue.put(("WARNING", "检测到 .ppt 文件，已自动切换为串行模式以避免 COM 冲突"))
            return self._run_sequential_files(all_files)

        # pptx 部分：并发或串行
        if use_concurrent and target_pptx:
            effective_workers = self._effective_workers_for(target_pptx)
            label = ".pptx 并发" if has_mixed else "并发"
            self.log_queue.put(("INFO", f"{label}转换已启用：最多 {effective_workers} 个子进程"))
            cancelled = self._run_subprocess_batch(target_pptx, effective_workers)
        elif target_pptx:
            cancelled = self._run_sequential_files(target_pptx)
        else:
            cancelled = False

        # 混合格式的 ppt 串行部分
        if has_mixed and not cancelled and not self.cancel_event.is_set():
            self.log_queue.put(("INFO", "开始串行转换 .ppt 文件"))
            cancelled = self._run_sequential_files(ppt_files)

        return cancelled

    def _resolve_max_workers(self) -> int:
        """解析并发子进程数（空值时自动）。"""
        if len(self.files) <= 1:
            return 1

        requested = self._parse_requested_max_workers()
        if requested is not None:
            return min(max(1, requested), len(self.files))

        try:
            cpu_count = mp.cpu_count()
        except NotImplementedError:
            cpu_count = 1

        auto_workers = max(1, min(_MAX_AUTO_WORKERS, cpu_count))
        return min(auto_workers, len(self.files))

    def _effective_workers_for(self, files: List[Tuple[int, Path]]) -> int:
        """计算某一批文件的有效并发上限。"""

        if not files:
            return 1
        return max(1, min(self._max_workers, len(files)))

    def _parse_requested_max_workers(self) -> Optional[int]:
        """读取用户配置的并发数（允许为空）。"""
        raw_value = self.params.get("max_workers")
        if raw_value is None:
            return None

        text = str(raw_value).strip()
        if not text:
            return None

        try:
            value = int(text)
        except (TypeError, ValueError):
            self.log_queue.put(("WARNING", f"并发进程数无效（{raw_value}），已回退自动模式"))
            return None

        if value <= 0:
            self.log_queue.put(("WARNING", f"并发进程数需大于 0（当前: {raw_value}），已回退自动模式"))
            return None

        return value

    def _run_sequential_files(self, files: List[Tuple[int, Path]]) -> bool:
        """串行执行指定文件列表（含子进程失败兜底）。"""

        for file_idx, file_path in files:
            if self.cancel_event.is_set():
                return True

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
                return True

            self._finalize_file_result(file_idx, file_path, success, detail)

        return False

    def _run_subprocess_batch(self, files: List[Tuple[int, Path]], max_workers: int) -> bool:
        """并发执行一批文件的子进程转换。"""

        pending = deque(files)
        running: Dict[int, _RunningSubprocessTask] = {}

        while pending or running:
            if self.cancel_event.is_set():
                self._terminate_running_tasks(running)
                return True

            while pending and len(running) < max_workers:
                file_idx, file_path = pending.popleft()
                self.log_queue.put(("INFO", f"正在处理: {file_path.name}"))

                try:
                    task = self._start_subprocess_task(file_idx, file_path)
                except Exception as exc:
                    error_message = f"子进程启动失败: {str(exc) or exc.__class__.__name__}"
                    self.log_queue.put(("ERROR", f"{file_path.name} 转换失败: {error_message}"))
                    self._finalize_file_result(file_idx, file_path, False, error_message)
                    continue

                running[file_idx] = task

            had_activity = False
            finished_indices: List[int] = []
            for file_idx, task in list(running.items()):
                drained_payload = self._drain_subprocess_messages(task.msg_queue, task.file_path, task.file_idx)
                if drained_payload is not None:
                    task.result_payload = drained_payload
                    had_activity = True

                if task.process.is_alive():
                    continue

                task.process.join()
                drained_payload = self._drain_subprocess_messages(task.msg_queue, task.file_path, task.file_idx)
                if drained_payload is not None:
                    task.result_payload = drained_payload

                success, detail = self._resolve_subprocess_result(task.process, task.result_payload)
                self._finalize_file_result(task.file_idx, task.file_path, success, detail)
                self._close_msg_queue(task.msg_queue)
                finished_indices.append(file_idx)
                had_activity = True

            for file_idx in finished_indices:
                running.pop(file_idx, None)

            if not had_activity:
                time.sleep(_POLL_INTERVAL_SEC)

        return False

    def _start_subprocess_task(self, file_idx: int, file_path: Path) -> _RunningSubprocessTask:
        """启动单文件子进程任务。"""

        msg_queue = self._mp_ctx.Queue()
        try:
            process = self._mp_ctx.Process(
                target=_convert_single_file_subprocess,
                args=(str(file_path), self.params, msg_queue),
                daemon=False,
            )
            process.start()
        except Exception:
            self._close_msg_queue(msg_queue)
            raise

        return _RunningSubprocessTask(
            file_idx=file_idx,
            file_path=file_path,
            process=process,
            msg_queue=msg_queue,
        )

    def _terminate_running_tasks(self, running: Dict[int, _RunningSubprocessTask]):
        """终止并回收所有在跑子进程。"""

        for task in running.values():
            try:
                task.process.terminate()
            except Exception:
                pass

        for task in running.values():
            try:
                task.process.join(timeout=5)
            except Exception:
                pass

            self._drain_subprocess_messages(task.msg_queue, task.file_path, task.file_idx)
            self._close_msg_queue(task.msg_queue)

        running.clear()

    def _close_msg_queue(self, msg_queue: Any):
        """关闭子进程消息队列。"""

        try:
            msg_queue.close()
            msg_queue.join_thread()
        except Exception:
            pass

    def _finalize_file_result(self, file_idx: int, file_path: Path, success: bool, detail: str):
        """统一落盘单文件结果并更新总体进度。"""

        if success:
            self._results.success_count += 1
        else:
            self._results.failed_count += 1
            self._results.failed_files.append((file_path, detail))
            if detail.startswith("子进程"):
                self.log_queue.put(("ERROR", f"{file_path.name} 转换失败: {detail}"))

        done_count = self._results.success_count + self._results.failed_count
        self._update_file_progress(file_idx, 1.0, f"已完成 {done_count}/{len(self.files)}")

    def _update_file_progress(self, file_idx: int, progress: float, status: str):
        """按文件粒度更新总体进度。"""

        if not self.files or file_idx not in self._file_progress:
            return

        normalized = max(0.0, min(1.0, float(progress)))
        prev = self._file_progress[file_idx]
        self._file_progress[file_idx] = max(prev, normalized)

        overall = sum(self._file_progress.values()) / len(self.files)
        self.progress_callback(overall, status)

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
            self._close_msg_queue(msg_queue)

        return self._resolve_subprocess_result(process, result_payload)

    @staticmethod
    def _resolve_subprocess_result(process: Any, result_payload: Optional[dict]) -> Tuple[bool, str]:
        """根据子进程退出状态与消息负载推断结果。"""

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
            status = f"正在处理: {file_path.name} ({int(current)}/{int(total)})"
            self._update_file_progress(file_idx, current / total, status)

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
            status = f"正在处理: {file_path.name} ({int(current)}/{int(total)})"
            self._update_file_progress(file_idx, current / total, status)

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

            status = f"正在处理: {file_path.name} ({int(current)}/{int(total)})"
            self._update_file_progress(file_idx, current / total, status)
            return None

        if kind == _SUBPROCESS_RESULT and len(message) >= 2 and isinstance(message[1], dict):
            return message[1]

        return None

    def get_results(self) -> ConversionResults:
        """获取转换结果。"""

        return self._results
