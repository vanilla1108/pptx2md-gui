"""pptx2md GUI 主题/配色集中配置。

目标：改配色只改这一处，避免在各个组件里到处搜硬编码颜色。
支持运行时浅色/深色模式切换，通过 observer 通知非 CTk 控件更新。
"""

from __future__ import annotations

import logging
from typing import Callable

import customtkinter as ctk

_LOGGER = logging.getLogger(__name__)

# -----------------------------------------------------------------------------
# 内部状态
# -----------------------------------------------------------------------------

_current_mode: str = "dark"  # "dark" | "light"
_listeners: list[Callable[[str], None]] = []

# CustomTkinter 内置主题名（或 .json 主题文件路径）
DEFAULT_COLOR_THEME = "dark-blue"

# 浅色模式窗口/容器基础背景色（暖奶油色调）
_WINDOW_BG_LIGHT = "#EDE8E0"

# -----------------------------------------------------------------------------
# 公共 API：模式管理
# -----------------------------------------------------------------------------


def apply_global_theme(initial_mode: str = "dark") -> None:
    """在创建控件前尽早调用，用于设置全局外观。

    参数:
        initial_mode: 初始外观模式（"dark" 或 "light"）。
    """
    global _current_mode
    _current_mode = initial_mode if initial_mode in ("dark", "light") else "dark"
    ctk.set_appearance_mode(_current_mode)
    ctk.set_default_color_theme(DEFAULT_COLOR_THEME)
    _patch_ctk_light_bg()


def _patch_ctk_light_bg() -> None:
    """将 CTk 内置主题中窗口和框架的浅色模式背景替换为暖色调。

    确保所有未显式设置 fg_color 的 CTkFrame 及根窗口在浅色模式下
    使用暖奶油色背景，与自定义语义化颜色方案保持一致。
    """
    theme_dict = ctk.ThemeManager.theme
    for widget_key in ("CTk", "CTkFrame", "CTkToplevel"):
        try:
            fg = theme_dict[widget_key]["fg_color"]
            if isinstance(fg, list) and len(fg) > 0:
                fg[0] = _WINDOW_BG_LIGHT
        except (KeyError, TypeError):
            _LOGGER.debug("跳过 %s 的浅色模式背景补丁", widget_key)
    try:
        top_fg = theme_dict["CTkFrame"]["top_fg_color"]
        if isinstance(top_fg, list) and len(top_fg) > 0:
            top_fg[0] = _WINDOW_BG_LIGHT
    except (KeyError, TypeError):
        _LOGGER.debug("跳过 CTkFrame top_fg_color 的浅色模式背景补丁")


def get_mode() -> str:
    """返回当前外观模式：'dark' 或 'light'。"""
    return _current_mode


def set_mode(mode: str) -> None:
    """设置外观模式并通知所有监听器。

    CTk 控件会自动跟随切换；非 CTk 控件需要通过监听器手动更新。
    """
    global _current_mode
    if mode not in ("dark", "light"):
        return
    _current_mode = mode
    ctk.set_appearance_mode(mode)
    for cb in list(_listeners):
        try:
            cb(mode)
        except Exception:
            _LOGGER.debug("主题监听器回调异常", exc_info=True)


def toggle_mode() -> str:
    """切换外观模式（dark <-> light），返回新模式。"""
    new_mode = "light" if _current_mode == "dark" else "dark"
    set_mode(new_mode)
    return new_mode


def register_on_mode_changed(callback: Callable[[str], None]) -> None:
    """注册外观模式变更监听器（用于需要手动更新的非 CTk 控件）。"""
    if callback not in _listeners:
        _listeners.append(callback)


def unregister_on_mode_changed(callback: Callable[[str], None]) -> None:
    """取消注册外观模式变更监听器。"""
    try:
        _listeners.remove(callback)
    except ValueError:
        pass


# -----------------------------------------------------------------------------
# 辅助函数
# -----------------------------------------------------------------------------


def paned_window_bg() -> str:
    """tk.PanedWindow 背景色：跟随当前 CustomTkinter 主题的 CTk 背景。"""
    mode_idx = 1 if ctk.get_appearance_mode() == "Dark" else 0
    return ctk.ThemeManager.theme["CTk"]["fg_color"][mode_idx]


def window_bg_pair() -> tuple[str, str]:
    """窗口基础背景色对 (light, dark)，可传给 CTk 控件的 fg_color。

    用于 tk.PanedWindow 直接子级的 CTkFrame（如 LogPanel），
    替代 ``fg_color="transparent"`` 以避免非 CTk 父控件下的颜色检测失败。
    必须在 :func:`apply_global_theme` 之后调用。
    """
    fg = ctk.ThemeManager.theme.get("CTk", {}).get("fg_color")
    if fg is None or len(fg) < 2:
        _LOGGER.warning("window_bg_pair() 在 apply_global_theme() 之前调用，使用回退值")
        return (_WINDOW_BG_LIGHT, "#242424")
    return (fg[0], fg[1])


def _resolve(pair: tuple[str, str]) -> str:
    """从 (light, dark) 颜色元组中按当前模式取值。"""
    return pair[1] if _current_mode == "dark" else pair[0]


# -----------------------------------------------------------------------------
# 语义化颜色：尽量用"用途"命名，而不是"具体色值"
# CTk 控件可直接使用 (light, dark) 元组，它们会自动跟随模式切换。
# -----------------------------------------------------------------------------

# 主色调 (Brand/Primary) - 蓝色在暖色背景上依然清爽
PRIMARY = ("#3B8ED0", "#3A7EBF")
PRIMARY_HOVER = ("#36719F", "#325E8C")

# 主色调按钮上的文字颜色（两种模式下都用白色保证对比度）
BTN_TEXT_ON_PRIMARY = ("#FFFFFF", "#FFFFFF")

# 面板/卡片背景（用于 ParamsPanel 分组、左侧预设卡片等）
# 浅色模式用暖白替代纯白，减少刺眼的高亮度
CARD_BG = ("#FAF8F5", "#2B2B2B")

# 通用浅表面（用于滚动列表等）
SURFACE_BG = ("#F5F2ED", "#232323")

# 更深的文本区域背景（用于日志文本框等）
SURFACE_BG_DEEP = ("#EFEBE4", "#1A1A1A")

# 边框颜色（暖灰色调，与背景协调）
BORDER_COLOR = ("#E0DBD3", "#404040")

# 次级/弱化文本颜色（暖灰，加深以满足 WCAG AA 4.5:1 对比度要求）
TEXT_MUTED = ("#665F58", "#A0A0A0")
# 主文本颜色（暖深灰替代近黑色，对比度从 ~18:1 降到 ~10:1，缓解视觉疲劳）
TEXT_PRIMARY = ("#3B3632", "#F3F4F6")

# 中性按钮（清空、浏览等基础按钮）
BTN_NEUTRAL_BG = ("#E5E0D8", "#404040")
BTN_NEUTRAL_HOVER = ("#D9D3CB", "#505050")
BTN_NEUTRAL_TEXT = ("#4A433C", "#E5E7EB")

# 危险动作 hover（用于"删除"这类：默认中性，悬停变红）
BTN_DANGER_HOVER = ("#FEE2E2", "#7F1D1D")
BTN_DANGER_TEXT_HOVER = ("#DC2626", "#FCA5A5")

# 危险动作主色（用于"取消转换"等：按钮本身即为红色）
BTN_DANGER_BG = ("#EF4444", "#DC2626")
BTN_DANGER_HOVER_DEEP = ("#DC2626", "#B91C1C")

# 拖拽区边框（默认/激活）—— 暖灰色调
DROP_BORDER = ("#D4CFC7", "#525252")
DROP_BORDER_ACTIVE = ("#3B8ED0", "#3A7EBF")

# 必填标记/徽章色
BADGE_TEXT = ("#EF4444", "#F87171")

# -----------------------------------------------------------------------------
# 日志级别颜色：(light, dark) 元组，None 表示使用默认文字颜色
# 浅色模式下使用更深的变体，保证在浅色背景上的可读性
# -----------------------------------------------------------------------------

_LOG_LEVEL_COLORS: dict[str, tuple[str, str] | None] = {
    "INFO": None,
    "WARNING": ("#B45309", "#F59E0B"),   # Amber 700 / 500
    "ERROR": ("#DC2626", "#EF4444"),     # Red 600 / 500
    "SUCCESS": ("#059669", "#10B981"),   # Emerald 600 / 500
    "DEBUG": ("#6B645D", "#888888"),     # 暖灰，加深以满足 WCAG AA 对比度
}


def log_level_color(level: str) -> str | None:
    """返回当前模式下指定日志级别的前景色。None 表示使用默认文字色。"""
    pair = _LOG_LEVEL_COLORS.get(level)
    if pair is None:
        return None
    return _resolve(pair)


# -----------------------------------------------------------------------------
# Tooltip 颜色（tk.Label 不支持 CTk 元组，通过函数按需解析）
# -----------------------------------------------------------------------------

_TOOLTIP_COLORS = {
    "bg": ("#FAF8F5", "#1F2937"),   # 暖白背景 / 深色背景
    "fg": ("#3B3632", "#F9FAFB"),   # 暖深灰文字 / 浅色文字
}


def tooltip_bg() -> str:
    """当前模式下的 Tooltip 背景色。"""
    return _resolve(_TOOLTIP_COLORS["bg"])


def tooltip_fg() -> str:
    """当前模式下的 Tooltip 前景色。"""
    return _resolve(_TOOLTIP_COLORS["fg"])
