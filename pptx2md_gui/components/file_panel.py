"""文件面板组件 - 主窗口左侧。"""

from pathlib import Path
from typing import Callable, List

import customtkinter as ctk

from .drop_zone import DropZone
from .. import theme


class FilePanel(ctk.CTkFrame):
    """左面板，包含文件拖放区、文件列表和预设选择。"""

    def __init__(
        self,
        master,
        on_files_changed: Callable[[List[Path]], None],
        **kwargs
    ):
        super().__init__(master, **kwargs)
        self.files: List[Path] = []
        self.on_files_changed = on_files_changed
        self._setup_ui()

    def _setup_ui(self):
        self.configure(fg_color="transparent")

        # 拖放区域
        self.drop_zone = DropZone(
            self,
            on_files_dropped=self.add_files,
            height=80,
        )
        self.drop_zone.pack(fill="x", padx=10, pady=(15, 10))

        # 文件列表标签
        list_label = ctk.CTkLabel(
            self,
            text="文件列表",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color=theme.TEXT_PRIMARY,
            anchor="w",
        )
        list_label.pack(fill="x", padx=12, pady=(5, 5))

        # 底部控制区 (优先布局，side="bottom" 确保始终显示)
        bottom_frame = ctk.CTkFrame(self, fg_color="transparent")
        bottom_frame.pack(side="bottom", fill="x", padx=10, pady=10)

        # 文件计数标签
        self.file_count_label = ctk.CTkLabel(
            bottom_frame,
            text="0 个文件",
            font=ctk.CTkFont(size=12),
            text_color=theme.TEXT_MUTED,
            anchor="w",
        )
        self.file_count_label.pack(side="left")

        # 清空按钮
        self.clear_btn = ctk.CTkButton(
            bottom_frame,
            text="清空",
            command=self.clear_files,
            width=60,
            height=24,
            font=ctk.CTkFont(size=12),
            fg_color=theme.BTN_NEUTRAL_BG,
            hover_color=theme.BTN_NEUTRAL_HOVER,
            text_color=theme.BTN_NEUTRAL_TEXT,
        )
        self.clear_btn.pack(side="right")

        # 文件列表框架（可滚动）
        self.file_list_frame = ctk.CTkScrollableFrame(
            self,
            height=200,
            fg_color=theme.SURFACE_BG,
            corner_radius=8,
            border_width=1,
            border_color=theme.BORDER_COLOR,
        )
        self.file_list_frame.pack(fill="both", expand=True, padx=10, pady=(0, 5))

    def add_files(self, paths: List[Path]):
        """添加文件到列表（自动去重，仅接受支持的格式）。"""
        from ..utils.validators import is_supported_file
        for path in paths:
            if not is_supported_file(path):
                continue
            if path not in self.files and path.exists():
                self.files.append(path)
        self._refresh_file_list()
        self.on_files_changed(self.files)

    def remove_file(self, path: Path):
        """从列表中移除单个文件。"""
        if path in self.files:
            self.files.remove(path)
        self._refresh_file_list()
        self.on_files_changed(self.files)

    def clear_files(self):
        """清空文件列表。"""
        self.files.clear()
        self._refresh_file_list()
        self.on_files_changed(self.files)

    def get_files(self) -> List[Path]:
        """获取当前文件列表。"""
        return self.files.copy()

    def has_ppt_files(self) -> bool:
        """检查文件列表中是否包含 .ppt 文件。"""
        return any(f.suffix.lower() == ".ppt" for f in self.files)

    def _refresh_file_list(self):
        """刷新文件列表显示。"""
        for widget in self.file_list_frame.winfo_children():
            widget.destroy()

        for path in self.files:
            self._create_file_item(path)

        self.file_count_label.configure(text=f"{len(self.files)} 个文件")

    def _create_file_item(self, path: Path):
        """在列表中创建文件项行。"""
        frame = ctk.CTkFrame(
            self.file_list_frame,
            fg_color="transparent",
            height=32,
        )
        frame.pack(fill="x", pady=2)
        frame.pack_propagate(False)

        # 文件图标（用文字模拟）
        icon_label = ctk.CTkLabel(
            frame,
            text="📄",
            font=ctk.CTkFont(size=14),
            width=24,
            anchor="center",
        )
        icon_label.pack(side="left", padx=(5, 0))

        # 文件名标签（超长截断）
        name = path.name
        if len(name) > 22:
            name = name[:19] + "..."

        label = ctk.CTkLabel(
            frame,
            text=name,
            font=ctk.CTkFont(size=13),
            text_color=theme.TEXT_PRIMARY,
            anchor="w",
        )
        label.pack(side="left", fill="x", expand=True, padx=5)

        # .ppt 文件的实验性标记
        if path.suffix.lower() == ".ppt":
            exp_label = ctk.CTkLabel(
                frame,
                text="⚠ 实验性",
                font=ctk.CTkFont(size=10),
                text_color=theme.TEXT_MUTED,
            )
            exp_label.pack(side="left", padx=(0, 5))

        # 删除按钮
        del_btn = ctk.CTkButton(
            frame,
            text="×",
            width=24,
            height=24,
            corner_radius=12,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color="transparent",
            hover_color=theme.BTN_DANGER_HOVER,
            text_color=theme.TEXT_MUTED,
            command=lambda p=path: self.remove_file(p),
        )
        del_btn.pack(side="right", padx=5)
        
        # Hover 效果：删除按钮在 hover 时变色，文字变色
        def on_enter(e):
            del_btn.configure(text_color=theme.BTN_DANGER_TEXT_HOVER)
        
        def on_leave(e):
            del_btn.configure(text_color=theme.TEXT_MUTED)

        del_btn.bind("<Enter>", on_enter)
        del_btn.bind("<Leave>", on_leave)
