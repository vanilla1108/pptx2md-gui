"""文件拖放区域组件。"""

import logging
from pathlib import Path
from typing import Callable, List

import customtkinter as ctk

from .. import theme

LOGGER = logging.getLogger(__name__)


class DropZone(ctk.CTkFrame):
    """PPTX 文件的拖放区域。"""

    IDLE_TEXT = "拖放 PPTX / PPT 文件到此处\n或点击选择"
    NO_DND_TEXT = "拖拽不可用\n请点击选择 PPTX / PPT 文件"
    DRAG_TEXT = "松开以添加文件"

    def __init__(
        self,
        master,
        on_files_dropped: Callable[[List[Path]], None],
        **kwargs
    ):
        super().__init__(master, **kwargs)
        self.on_files_dropped = on_files_dropped
        self.dnd_available = False
        self._setup_ui()
        self._setup_dnd()
        if not self.dnd_available:
            self._set_idle_state(no_dnd=True)

    def _setup_ui(self):
        self.configure(
            corner_radius=12,
            border_width=2,
            border_color=theme.DROP_BORDER,
            fg_color=theme.SURFACE_BG,
        )

        # 使用内部 frame 居中，图标与文字上下布局，避免文字被横向挤压
        content_frame = ctk.CTkFrame(self, fg_color="transparent")
        content_frame.place(relx=0.5, rely=0.5, anchor="center")

        # 图标（Unicode）
        self.icon_label = ctk.CTkLabel(
            content_frame,
            text="📂",
            font=ctk.CTkFont(size=24),
            text_color=theme.TEXT_MUTED,
        )
        self.icon_label.pack(pady=(0, 4))

        self.label = ctk.CTkLabel(
            content_frame,
            text=self.IDLE_TEXT,
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color=theme.TEXT_MUTED,
            justify="center",
            anchor="center",
        )
        self.label.pack()

        self.bind("<Button-1>", self._on_click)
        self.label.bind("<Button-1>", self._on_click)
        self.icon_label.bind("<Button-1>", self._on_click)

    def _setup_dnd(self):
        """使用 tkinterdnd2 设置拖放功能。"""
        dnd_enabled = bool(getattr(self.winfo_toplevel(), "dnd_available", False))
        if not dnd_enabled:
            LOGGER.warning("当前窗口未启用 tkinterdnd2，拖拽功能不可用")
            return

        try:
            self.drop_target_register("DND_Files")
            self.dnd_bind("<<Drop>>", self._on_drop)
            self.dnd_bind("<<DragEnter>>", self._on_drag_enter)
            self.dnd_bind("<<DragLeave>>", self._on_drag_leave)
            self.dnd_available = True
        except Exception as exc:
            LOGGER.warning("注册拖拽事件失败，拖拽功能不可用: %s", exc)

    def _on_click(self, event=None):
        from tkinter import filedialog
        files = filedialog.askopenfilenames(
            title="选择 PowerPoint 文件",
            filetypes=[
                ("PowerPoint 文件", "*.pptx *.ppt"),
                ("PPTX 文件", "*.pptx"),
                ("PPT 文件 (实验性)", "*.ppt"),
                ("所有文件", "*.*"),
            ],
        )
        if files:
            self.on_files_dropped([Path(f) for f in files])

    def _on_drop(self, event):
        files = self._parse_drop_data(event.data)
        from ..utils.validators import is_supported_file
        supported_files = [f for f in files if is_supported_file(f)]
        if supported_files:
            self.on_files_dropped(supported_files)
        self._on_drag_leave(event)

    def _on_drag_enter(self, event):
        self.configure(border_color=theme.DROP_BORDER_ACTIVE)
        self.label.configure(text=self.DRAG_TEXT, text_color=theme.PRIMARY)
        self.icon_label.configure(text_color=theme.PRIMARY)

    def _on_drag_leave(self, event):
        self._set_idle_state(no_dnd=not self.dnd_available)

    def _set_idle_state(self, no_dnd: bool = False):
        """恢复空闲态文案与样式。"""
        self.configure(border_color=theme.DROP_BORDER)
        self.label.configure(
            text=self.NO_DND_TEXT if no_dnd else self.IDLE_TEXT,
            text_color=theme.TEXT_MUTED,
        )
        self.icon_label.configure(text_color=theme.TEXT_MUTED)

    def _parse_drop_data(self, data: str) -> List[Path]:
        """解析 tkinterdnd2 的拖放文件数据。"""
        files = []
        if data.startswith("{"):
            for item in data.split("} {"):
                item = item.strip("{}")
                if item:
                    files.append(Path(item))
        else:
            for item in data.split():
                if item:
                    files.append(Path(item))
        return files
