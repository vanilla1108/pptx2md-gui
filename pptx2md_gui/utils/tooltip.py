"""CustomTkinter 控件的工具提示实现。"""

import tkinter as tk
from typing import Optional

from .. import theme


class Tooltip:
    """鼠标悬停在控件上时显示的工具提示。"""

    def __init__(self, widget: tk.Widget, text: str, delay: int = 500):
        """初始化工具提示。

        参数:
            widget: 要附加工具提示的控件。
            text: 工具提示文本。
            delay: 显示工具提示前的延迟时间（毫秒）。
        """
        self.widget = widget
        self.text = text
        self.delay = delay
        self._tipwindow: Optional[tk.Toplevel] = None
        self._id_after: Optional[str] = None

        self.widget.bind("<Enter>", self._on_enter)
        self.widget.bind("<Leave>", self._on_leave)

    def _on_enter(self, event=None):
        self._id_after = self.widget.after(self.delay, self._show)

    def _on_leave(self, event=None):
        if self._id_after:
            self.widget.after_cancel(self._id_after)
            self._id_after = None
        self._hide()

    def _show(self):
        if self._tipwindow:
            return

        x = self.widget.winfo_rootx() + 20
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 5

        self._tipwindow = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")

        label = tk.Label(
            tw,
            text=self.text,
            justify="left",
            background=theme.tooltip_bg(),
            foreground=theme.tooltip_fg(),
            relief="solid",
            borderwidth=1,
            font=("Segoe UI", 9),
            padx=6,
            pady=3,
        )
        label.pack()

    def _hide(self):
        if self._tipwindow:
            self._tipwindow.destroy()
            self._tipwindow = None

    def update_text(self, text: str):
        """更新工具提示文本。"""
        self.text = text
