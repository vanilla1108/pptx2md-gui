"""pptx2md GUI 主应用窗口。"""

import logging
import queue
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import messagebox
from typing import List

import customtkinter as ctk
from tkinterdnd2 import TkinterDnD

import pptx2md
from pptx2md.log import setup_logging

from . import theme
from .components.file_panel import FilePanel
from .components.log_panel import LogPanel
from .components.params_panel import ParamsPanel
from .services.converter import ConversionResults, ConversionWorker, QueueLogHandler
from .services.preset_manager import PresetManager

LOG_POLL_INTERVAL_MS = 100
MIN_LOG_PANEL_HEIGHT = 150  # 日志面板最小高度（px），防止被完全折叠
SASH_HIT_WIDTH = 10  # 分隔条可拖拽命中宽度（px）
LOGGER = logging.getLogger(__name__)


class DnDCompatibleCTk(ctk.CTk):
    """支持 tkinterdnd2 拖拽的 CTk 根窗口。"""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.dnd_available = False
        self._enable_dnd()

    def _enable_dnd(self):
        """在根窗口上启用 tkdnd 拖拽包。"""
        try:
            self.TkdndVersion = TkinterDnD._require(self)
            self.dnd_available = True
        except Exception as exc:
            LOGGER.warning("初始化拖拽支持失败，将退化为点击选择: %s", exc)


class App(DnDCompatibleCTk):
    """主应用窗口。"""

    def __init__(self):
        super().__init__()

        # 窗口配置
        self.title("pptx2md - PPT 转 Markdown 工具")
        self.geometry("1000x700")
        self.minsize(900, 650)

        # 状态
        self._log_queue = queue.Queue()
        self._cancel_event = threading.Event()
        self._worker = None
        self._preset_manager = PresetManager()
        self._ppt_warned = False  # 去重标志：避免重复输出 PPT 环境检测日志

        # 全局主题（配色集中在 pptx2md_gui/theme.py）
        # 从预设管理器加载上次使用的外观模式
        saved_mode = self._preset_manager.get_appearance_mode()
        theme.apply_global_theme(initial_mode=saved_mode)

        # 为 GUI 设置带队列处理器的日志系统
        self._log_handler = QueueLogHandler(self._log_queue)
        setup_logging(compat_tqdm=False, external_handlers=[self._log_handler])
        if getattr(pptx2md, "__file__", None) is None:
            LOGGER.warning("无法定位 pptx2md 模块路径，转换功能可能异常")

        # 构建界面
        self._setup_ui()

        # 注册外观模式变更监听（更新 tk.PanedWindow 背景色 + 持久化偏好）
        self._mode_changed_cb = self._on_mode_changed
        theme.register_on_mode_changed(self._mode_changed_cb)

        # 加载上次使用的预设
        self._load_last_preset()

        # 启动日志轮询
        self._poll_log_queue()

    def _setup_ui(self):
        """构建主窗口布局。"""
        # 主容器（单行单列，垂直 PanedWindow 占满）
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # 垂直分隔窗口：上方面板区 + 下方日志区，用户可拖拽调节高度
        # 使用 tk.PanedWindow（非 ttk）以支持 bg/opaqueresize/paneconfigure(minsize)
        bg = theme.paned_window_bg()
        self._vpaned = tk.PanedWindow(
            self,
            orient=tk.VERTICAL,
            bg=bg,
            # 显式指定普通光标，避免在 Windows + Tk 下出现 sashcursor “泄漏”到整块区域的情况
            # （升级样式后更容易在 PanedWindow 背景上触发，从而看起来全局都是双向箭头）。
            cursor="arrow",
            sashwidth=SASH_HIT_WIDTH,
            sashrelief="flat",
            # Windows 下该选项会在部分 Tk 版本中“泄漏”到整块区域，导致默认光标异常
            sashcursor="arrow",
            opaqueresize=False,
            borderwidth=0,
        )
        self._vpaned.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        self._bind_paned_sash_cursor(self._vpaned, "sb_v_double_arrow")

        # ── 上半部分：含左右面板的水平分隔窗口 ──
        self._hpaned = tk.PanedWindow(
            self._vpaned,
            orient=tk.HORIZONTAL,
            bg=bg,
            cursor="arrow",
            sashwidth=SASH_HIT_WIDTH,
            sashrelief="flat",
            sashcursor="arrow",
            opaqueresize=False,
            borderwidth=0,
        )
        self._bind_paned_sash_cursor(self._hpaned, "sb_h_double_arrow")

        # 左面板 - 文件选择（网格布局：FilePanel 可伸缩，预设区域固定高度）
        # bg_color 显式指定：父控件是 tk.PanedWindow（非 CTk），CTk 无法自动检测其背景色，
        # 导致圆角外露区域使用错误的默认色（白色），在深色模式下出现白色三角形。
        _paned_bg = theme.window_bg_pair()
        left_container = ctk.CTkFrame(self._hpaned, width=320, bg_color=_paned_bg)
        left_container.configure(cursor="arrow")
        left_container.pack_propagate(False)
        left_container.grid_propagate(False)
        left_container.grid_rowconfigure(0, weight=1)
        left_container.grid_rowconfigure(1, weight=0)
        left_container.grid_columnconfigure(0, weight=1)

        self.file_panel = FilePanel(
            left_container,
            on_files_changed=self._on_files_changed,
        )
        self.file_panel.grid(row=0, column=0, sticky="nsew")

        # 预设区域（左面板内部，底部固定高度）
        self._create_preset_section(left_container)

        self._hpaned.add(left_container, stretch="never", minsize=280, sticky="nsew")

        # 右面板 - 参数设置（可滚动）
        right_container = ctk.CTkFrame(self._hpaned, bg_color=_paned_bg)
        right_container.configure(cursor="arrow")

        self.params_panel = ParamsPanel(right_container)
        self.params_panel.pack(fill="both", expand=True, padx=5, pady=5)

        self._hpaned.add(right_container, stretch="always", sticky="nsew")

        self._vpaned.add(self._hpaned, stretch="always", sticky="nsew")

        # ── 下半部分：日志面板（可拖拽调节高度） ──
        self.log_panel = LogPanel(
            self._vpaned,
            on_start=self._on_start_conversion,
            on_cancel=self._on_cancel_conversion,
        )
        self.log_panel.configure(cursor="arrow")
        self._vpaned.add(
            self.log_panel,
            stretch="never",
            minsize=MIN_LOG_PANEL_HEIGHT,
            sticky="nsew",
        )

        # 初始按钮状态
        self.log_panel.set_start_enabled(False)

    @staticmethod
    def _bind_paned_sash_cursor(paned: tk.PanedWindow, sash_cursor: str):
        """仅在鼠标位于 sash/handle 时显示双向箭头，其他区域保持普通箭头。"""

        def _pointer_on_self() -> bool:
            try:
                x_root = paned.winfo_pointerx()
                y_root = paned.winfo_pointery()
                target = paned.winfo_containing(x_root, y_root)
                return target == paned
            except tk.TclError:
                return False

        def _is_on_sash(x: int, y: int) -> bool:
            try:
                hit = paned.identify(x, y)
            except tk.TclError:
                return False

            if not hit:
                return False

            if isinstance(hit, (list, tuple)):
                parts = [str(item).lower() for item in hit]
                return "sash" in parts or "handle" in parts

            hit_text = str(hit).lower()
            return "sash" in hit_text or "handle" in hit_text

        def _set_cursor(event=None):
            try:
                x = paned.winfo_pointerx() - paned.winfo_rootx()
                y = paned.winfo_pointery() - paned.winfo_rooty()

                if not _pointer_on_self():
                    paned.configure(cursor="arrow")
                    return

                paned.configure(cursor=sash_cursor if _is_on_sash(x, y) else "arrow")
            except tk.TclError:
                return

        paned.bind("<Enter>", _set_cursor, add="+")
        paned.bind("<Motion>", _set_cursor, add="+")
        paned.bind("<ButtonPress-1>", _set_cursor, add="+")
        paned.bind("<B1-Motion>", _set_cursor, add="+")
        paned.bind("<ButtonRelease-1>", _set_cursor, add="+")
        paned.bind("<Leave>", lambda _event: paned.configure(cursor="arrow"), add="+")

    def _create_preset_section(self, parent):
        """在左面板中创建预设选择区域。"""
        preset_frame = ctk.CTkFrame(
            parent, 
            fg_color=theme.CARD_BG, 
            corner_radius=8,
            border_width=1,
            border_color=theme.BORDER_COLOR,
        )
        preset_frame.grid(row=1, column=0, sticky="ew", padx=10, pady=(5, 10))
        
        # 使用 Grid 布局以获得更好的对齐控制
        preset_frame.grid_columnconfigure(1, weight=1)

        # 第一行：标签 + 下拉菜单
        ctk.CTkLabel(
            preset_frame,
            text="配置预设:",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color=theme.TEXT_PRIMARY,
        ).grid(row=0, column=0, padx=(12, 5), pady=12, sticky="w")

        preset_names = self._preset_manager.get_preset_names()
        self._preset_var = ctk.StringVar(value=self._preset_manager.get_last_used())
        self._preset_dropdown = ctk.CTkOptionMenu(
            preset_frame,
            values=preset_names if preset_names else ["默认配置"],
            variable=self._preset_var,
            command=self._on_preset_selected,
            fg_color=theme.SURFACE_BG_DEEP,
            button_color=theme.SURFACE_BG_DEEP,
            button_hover_color=theme.BTN_NEUTRAL_HOVER,
            text_color=theme.TEXT_PRIMARY,
            height=28,
            anchor="w",  # 文字左对齐
        )
        self._preset_dropdown.grid(row=0, column=1, padx=(0, 12), pady=12, sticky="ew")

        # 第二行：按钮组
        btn_frame = ctk.CTkFrame(preset_frame, fg_color="transparent")
        btn_frame.grid(row=1, column=0, columnspan=2, sticky="ew", padx=12, pady=(0, 12))

        ctk.CTkButton(
            btn_frame,
            text="保存当前配置",
            height=28,
            fg_color=theme.BTN_NEUTRAL_BG,
            hover_color=theme.BTN_NEUTRAL_HOVER,
            text_color=theme.BTN_NEUTRAL_TEXT,
            command=self._save_preset,
        ).pack(side="left", fill="x", expand=True, padx=(0, 5))

        ctk.CTkButton(
            btn_frame,
            text="删除",
            width=60,
            height=28,
            fg_color=theme.BTN_NEUTRAL_BG,
            hover_color=theme.BTN_DANGER_HOVER,
            text_color=theme.BTN_NEUTRAL_TEXT,
            command=self._delete_preset,
        ).pack(side="left")

    def _on_mode_changed(self, mode: str):
        """外观模式变更回调：更新非 CTk 控件配色并持久化偏好。"""
        bg = theme.paned_window_bg()
        self._vpaned.configure(bg=bg)
        self._hpaned.configure(bg=bg)
        self._preset_manager.set_appearance_mode(mode)

    def destroy(self):
        """销毁窗口前清理模块级监听器，避免回调泄漏。"""
        cb = getattr(self, "_mode_changed_cb", None)
        if cb is not None:
            theme.unregister_on_mode_changed(cb)
        super().destroy()

    def _on_files_changed(self, files: List[Path]):
        """处理文件列表变更。"""
        has_files = len(files) > 0
        self.log_panel.set_start_enabled(has_files)

        # PPT 参数灰置联动
        has_ppt = self.file_panel.has_ppt_files()
        self.params_panel.set_ppt_group_enabled(has_ppt)
        if has_ppt and not self._ppt_warned:
            self._ppt_warned = True
            from pptx2md.ppt_legacy import check_environment
            env_ok, env_reason = check_environment(strict=False)
            if env_ok:
                self.log_panel.log("INFO", "PPT 格式转换为实验性功能，当前环境满足要求")
            else:
                self.log_panel.log("WARNING", f"PPT 格式转换为实验性功能，{env_reason}")
        elif not has_ppt:
            self._ppt_warned = False

    def _on_start_conversion(self):
        """启动转换流程。"""
        files = self.file_panel.get_files()
        if not files:
            messagebox.showwarning("提示", "请先选择要转换的文件")
            return

        params = self.params_panel.get_params()

        # 验证输出目录
        if not params.get("output_dir"):
            # 使用第一个文件的父目录作为默认值
            params["output_dir"] = str(files[0].parent)

        # 重置状态
        self._cancel_event.clear()
        self.log_panel.reset_progress()
        self.log_panel.set_converting(True)

        # 创建并启动工作线程
        self._worker = ConversionWorker(
            files=files,
            params=params,
            log_queue=self._log_queue,
            progress_callback=self._on_progress_update,
            cancel_event=self._cancel_event,
            on_complete=self._on_conversion_complete,
        )
        self._worker.start()

    def _on_cancel_conversion(self):
        """取消当前转换。"""
        if self._worker and self._worker.is_alive():
            self._cancel_event.set()
            self.log_panel.log("WARNING", "正在取消转换...")

    def _on_progress_update(self, value: float, status: str):
        """处理来自工作线程的进度更新（在工作线程中调用）。"""
        # 在主线程上调度 UI 更新
        self.after(0, lambda: self.log_panel.set_progress(value, status))

    def _on_conversion_complete(self, results: ConversionResults):
        """处理转换完成事件（在工作线程中调用）。"""
        self.after(0, lambda: self._handle_completion(results))

    def _handle_completion(self, results: ConversionResults):
        """在主线程上处理转换完成事件。"""
        self.log_panel.set_converting(False)
        self.log_panel.set_progress(1.0, "完成")

        if results.failed_count > 0:
            failed_names = ", ".join(f.name for f, _ in results.failed_files)
            messagebox.showwarning(
                "部分转换失败",
                f"成功: {results.success_count}, 失败: {results.failed_count}\n"
                f"失败文件: {failed_names}",
            )

    def _poll_log_queue(self):
        """轮询日志队列并更新界面。"""
        try:
            while True:
                level, message = self._log_queue.get_nowait()
                self.log_panel.log(level, message)
        except queue.Empty:
            pass

        self.after(LOG_POLL_INTERVAL_MS, self._poll_log_queue)

    def _on_preset_selected(self, name: str):
        """处理预设选择。"""
        preset = self._preset_manager.get_preset(name)
        if preset:
            self.params_panel.set_params(preset)
            self._preset_manager.set_last_used(name)

    def _save_preset(self):
        """将当前参数保存为预设。"""
        name = self._preset_var.get().strip() or "默认配置"

        params = self.params_panel.get_params()
        self._preset_manager.save_preset(name, params)
        self._preset_manager.set_last_used(name)
        self._refresh_preset_dropdown()
        self._preset_var.set(name)

    def _delete_preset(self):
        """删除选中的预设。"""
        name = self._preset_var.get()
        if name == "默认配置":
            messagebox.showinfo("提示", "不能删除默认配置")
            return

        if messagebox.askyesno("确认", f"确定删除预设 \"{name}\"？"):
            if self._preset_manager.delete_preset(name):
                self._refresh_preset_dropdown()
                last = self._preset_manager.get_last_used()
                self._preset_var.set(last)
                self._on_preset_selected(last)

    def _refresh_preset_dropdown(self):
        """刷新预设下拉菜单的选项。"""
        names = self._preset_manager.get_preset_names()
        self._preset_dropdown.configure(values=names if names else ["默认配置"])

    def _load_last_preset(self):
        """启动时加载上次使用的预设。"""
        last = self._preset_manager.get_last_used()
        preset = self._preset_manager.get_preset(last)
        if preset:
            self.params_panel.set_params(preset)
