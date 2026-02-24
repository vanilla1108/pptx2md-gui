"""GUI 字段的输入验证器。"""

from pathlib import Path
from typing import Optional, Tuple


def validate_integer(value: str, min_val: int = 0, max_val: int = 99999) -> Tuple[bool, Optional[str]]:
    """验证整数输入字符串。

    参数:
        value: 要验证的字符串值。
        min_val: 允许的最小值。
        max_val: 允许的最大值。

    返回:
        (是否有效, 错误消息) 元组。
    """
    if not value or value.strip() == "":
        return True, None

    try:
        num = int(value)
        if num < min_val:
            return False, f"值不能小于 {min_val}"
        if num > max_val:
            return False, f"值不能大于 {max_val}"
        return True, None
    except ValueError:
        return False, "请输入有效的整数"


def validate_page_number(value: str) -> Tuple[bool, Optional[str]]:
    """验证页码输入。

    参数:
        value: 要验证的字符串值。

    返回:
        (是否有效, 错误消息) 元组。
    """
    if not value or value.strip() == "":
        return True, None

    return validate_integer(value, min_val=1)


def validate_path(value: str, must_exist: bool = False) -> Tuple[bool, Optional[str]]:
    """验证文件路径。

    参数:
        value: 要验证的路径字符串。
        must_exist: 路径是否必须存在。

    返回:
        (是否有效, 错误消息) 元组。
    """
    if not value or value.strip() == "":
        return True, None

    try:
        path = Path(value)
        if must_exist and not path.exists():
            return False, f"路径不存在: {value}"
        return True, None
    except (OSError, ValueError):
        return False, f"无效的路径: {value}"


# ---------------------------------------------------------------------------
# 文件扩展名验证
# ---------------------------------------------------------------------------

SUPPORTED_EXTENSIONS = {".pptx", ".ppt"}


def is_supported_file(path: Path) -> bool:
    """检查文件扩展名是否受支持。"""
    return path.suffix.lower() in SUPPORTED_EXTENSIONS


def has_ppt_in_list(files: list[Path]) -> bool:
    """检查文件列表中是否包含 .ppt 文件。"""
    return any(f.suffix.lower() == ".ppt" for f in files)
