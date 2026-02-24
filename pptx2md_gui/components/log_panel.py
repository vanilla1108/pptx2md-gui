"""日志面板组件 - 主窗口底部。"""

import math
import queue
from datetime import datetime
from typing import Callable, Optional

import customtkinter as ctk
from PIL import Image, ImageDraw

from .. import theme

# ---------------------------------------------------------------------------
# 模式切换图标绘制（Pillow 矢量 + 4x 超采样抗锯齿）
# ---------------------------------------------------------------------------

_ICON_PX = 16  # 图标显示尺寸
_SS = 4  # 超采样倍数


def _draw_sun(color: str) -> Image.Image:
    """绘制太阳图标：圆形 + 8 条射线。"""
    s = _ICON_PX * _SS
    img = Image.new("RGBA", (s, s), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    c = s / 2
    body_r = s * 0.18
    draw.ellipse([c - body_r, c - body_r, c + body_r, c + body_r], fill=color)

    ray_in = body_r + s * 0.07
    ray_out = c - s * 0.06
    w = max(1, round(s * 0.07))
    for i in range(8):
        angle = i * math.pi / 4
        cos_a, sin_a = math.cos(angle), math.sin(angle)
        draw.line(
            [(c + ray_in * cos_a, c + ray_in * sin_a),
             (c + ray_out * cos_a, c + ray_out * sin_a)],
            fill=color,
            width=w,
        )

    return img.resize((_ICON_PX, _ICON_PX), Image.LANCZOS)


def _draw_crescent(color: str) -> Image.Image:
    """绘制月牙图标：圆形减去偏移圆形。"""
    s = _ICON_PX * _SS

    # 用灰度蒙版绘制月牙：主圆 + 偏移挖去
    mask = Image.new("L", (s, s), 0)
    draw = ImageDraw.Draw(mask)

    pad = s * 0.12
    draw.ellipse([pad, pad, s - pad, s - pad], fill=255)

    # 向右上方偏移的圆形，"咬掉"一块形成月牙
    offset_x, offset_y = s * 0.22, s * -0.08
    cut_r = (s - 2 * pad) * 0.43
    cx, cy = s / 2 + offset_x, s / 2 + offset_y
    draw.ellipse([cx - cut_r, cy - cut_r, cx + cut_r, cy + cut_r], fill=0)

    # 颜色层 + 蒙版 → 带透明度的月牙
    colored = Image.new("RGB", (s, s), color)
    colored.putalpha(mask)

    return colored.resize((_ICON_PX, _ICON_PX), Image.LANCZOS)


def _create_mode_toggle_icon() -> ctk.CTkImage:
    """创建模式切换图标。

    CTkImage 根据当前外观模式自动显示对应图标：
    - 浅色模式 → 月牙（深色，提示"点击切到深色"）
    - 深色模式 → 太阳（浅色，提示"点击切到浅色"）
    """
    return ctk.CTkImage(
        light_image=_draw_crescent(theme.BTN_NEUTRAL_TEXT[0]),
        dark_image=_draw_sun(theme.BTN_NEUTRAL_TEXT[1]),
        size=(_ICON_PX, _ICON_PX),
    )


class LogPanel(ctk.CTkFrame):
    """底部面板，包含日志显示、进度条和操作按钮。"""

    def __init__(
        self,
        master,
        on_start: Callable[[], None],
        on_cancel: Callable[[], None],
        **kwargs
    ):
        # LogPanel 的父控件是 tk.PanedWindow（非 CTk 控件），CTk 无法自动检测其背景色。
        # 显式设置 bg_color 避免圆角外露区域使用错误的默认色（白色三角形）。
        kwargs.setdefault("bg_color", theme.window_bg_pair())
        super().__init__(master, **kwargs)
        self.on_start = on_start
        self.on_cancel = on_cancel
        self.log_queue: Optional[queue.Queue] = None
        self._setup_ui()

        # 注册主题变更监听，用于更新日志标签颜色和切换按钮图标
        self._mode_changed_cb = self._on_mode_changed
        theme.register_on_mode_changed(self._mode_changed_cb)

    def _setup_ui(self):
        # 使用显式背景色替代 "transparent"，避免 tk.PanedWindow（非 CTk 控件）
        # 作为父级时 CTk 的颜色检测失败，导致模式切换后外边框颜色不正确
        self.configure(fg_color=theme.window_bg_pair())

        # 1. 底部按钮区域 (优先布局，确保始终可见)
        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(side="bottom", fill="x", padx=10, pady=(5, 10))

        # 间隔器，将按钮推到右侧
        ctk.CTkLabel(btn_frame, text="").pack(side="left", fill="x", expand=True)

        self.clear_btn = ctk.CTkButton(
            btn_frame,
            text="清空日志",
            width=80,
            fg_color=theme.BTN_NEUTRAL_BG,
            hover_color=theme.BTN_NEUTRAL_HOVER,
            text_color=theme.BTN_NEUTRAL_TEXT,
            command=self.clear_log,
        )
        self.clear_btn.pack(side="left", padx=5)

        self.start_btn = ctk.CTkButton(
            btn_frame,
            text="开始转换",
            width=120,
            fg_color=theme.PRIMARY,
            hover_color=theme.PRIMARY_HOVER,
            text_color=theme.BTN_TEXT_ON_PRIMARY,
            font=ctk.CTkFont(weight="bold"),
            command=self.on_start,
        )
        self.start_btn.pack(side="left", padx=5)

        # 2. 进度条区域 (位于按钮上方)
        progress_frame = ctk.CTkFrame(self, fg_color="transparent")
        progress_frame.pack(side="bottom", fill="x", padx=10, pady=5)

        ctk.CTkLabel(
            progress_frame,
            text="进度:",
            text_color=theme.TEXT_PRIMARY
        ).pack(side="left")

        self.progress_bar = ctk.CTkProgressBar(
            progress_frame,
            width=300,
            progress_color=theme.PRIMARY,
        )
        self.progress_bar.pack(side="left", padx=10, fill="x", expand=True)
        self.progress_bar.set(0)

        self.progress_label = ctk.CTkLabel(
            progress_frame, text="0%", width=50, anchor="e", text_color=theme.TEXT_PRIMARY
        )
        self.progress_label.pack(side="left")

        self.status_label = ctk.CTkLabel(
            progress_frame, text="就绪", text_color=theme.TEXT_MUTED
        )
        self.status_label.pack(side="left", padx=10)

        # 3. 顶部标题区域
        header_frame = ctk.CTkFrame(self, fg_color="transparent")
        header_frame.pack(side="top", fill="x", padx=10, pady=(5, 2))

        log_label = ctk.CTkLabel(
            header_frame,
            text="转换日志",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color=theme.TEXT_PRIMARY,
            anchor="w",
        )
        log_label.pack(side="left")

        # 外观模式切换按钮（标题栏右侧）
        # CTkImage 自动根据当前模式显示太阳/月牙，无需手动切换
        self._mode_icon = _create_mode_toggle_icon()
        self._mode_btn = ctk.CTkButton(
            header_frame,
            image=self._mode_icon,
            text="",
            width=28,
            height=28,
            corner_radius=6,
            fg_color=theme.BTN_NEUTRAL_BG,
            hover_color=theme.BTN_NEUTRAL_HOVER,
            command=self._on_toggle_mode,
        )
        self._mode_btn.pack(side="right")

        # 4. 中间日志文本区域 (占用剩余空间)
        self.log_text = ctk.CTkTextbox(
            self,
            height=120,
            font=ctk.CTkFont(family="Consolas", size=12),
            fg_color=theme.SURFACE_BG_DEEP,
            text_color=theme.TEXT_PRIMARY,
            corner_radius=8,
            border_width=1,
            border_color=theme.BORDER_COLOR,
        )
        self.log_text.pack(side="top", fill="both", expand=True, padx=10, pady=(0, 5))
        self.log_text.configure(state="disabled")

        # 记录默认按钮样式，用于状态切换时恢复
        self._start_btn_default_fg = theme.PRIMARY
        self._start_btn_default_hover = theme.PRIMARY_HOVER
        self._converting = False

    def _on_toggle_mode(self):
        """处理外观模式切换按钮点击。"""
        theme.toggle_mode()

    def _on_mode_changed(self, _mode: str):
        """主题变更回调：更新已有日志行的颜色标签。

        模式切换按钮的图标由 CTkImage 自动切换（light_image/dark_image），无需手动更新。
        """
        self._reapply_log_tag_colors()

    def _reapply_log_tag_colors(self):
        """用当前模式的颜色重新设置已有日志行的 tag foreground。"""
        for level in ("WARNING", "ERROR", "SUCCESS", "DEBUG"):
            color = theme.log_level_color(level)
            if color:
                try:
                    self.log_text.tag_config(f"level_{level}", foreground=color)
                except Exception:
                    pass

    def log(self, level: str, message: str):
        """添加带时间戳和级别的日志消息。"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        formatted = f"[{timestamp}] [{level}] {message}\n"

        self.log_text.configure(state="normal")
        self.log_text.insert("end", formatted)

        # 如果不是默认颜色，则应用颜色标签
        color = theme.log_level_color(level)
        if color:
            start_idx = self.log_text.index("end-2l linestart")
            end_idx = self.log_text.index("end-1l lineend")
            tag_name = f"level_{level}"
            self.log_text.tag_config(tag_name, foreground=color)
            self.log_text.tag_add(tag_name, start_idx, end_idx)

        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def clear_log(self):
        """清空所有日志消息。"""
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

    def set_progress(self, value: float, status: str = ""):
        """设置进度条值（0.0 到 1.0）和状态文本。"""
        self.progress_bar.set(value)
        self.progress_label.configure(text=f"{int(value * 100)}%")
        if status:
            self.status_label.configure(text=status)

    def set_converting(self, converting: bool):
        """根据转换状态切换按钮的文本、样式和回调。"""
        self._converting = converting
        if converting:
            self.start_btn.configure(
                text="取消转换",
                fg_color=theme.BTN_DANGER_BG,
                hover_color=theme.BTN_DANGER_HOVER_DEEP,
                command=self.on_cancel,
                state="normal",
            )
        else:
            self.start_btn.configure(
                text="开始转换",
                fg_color=self._start_btn_default_fg,
                hover_color=self._start_btn_default_hover,
                command=self.on_start,
                state="normal",
            )

    def set_start_enabled(self, enabled: bool):
        """启用或禁用开始按钮（转换进行中时不受影响）。"""
        if self._converting:
            return
        self.start_btn.configure(state="normal" if enabled else "disabled")

    def reset_progress(self):
        """重置进度条和状态。"""
        self.progress_bar.set(0)
        self.progress_label.configure(text="0%")
        self.status_label.configure(text="就绪")

    def destroy(self):
        """销毁前取消主题监听，避免残留回调引用已销毁控件。"""
        cb = getattr(self, "_mode_changed_cb", None)
        if cb is not None:
            theme.unregister_on_mode_changed(cb)
        super().destroy()
