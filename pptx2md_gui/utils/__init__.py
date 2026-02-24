"""pptx2md GUI 的工具函数。"""

from .tooltip import Tooltip
from .validators import validate_integer, validate_page_number, validate_path

__all__ = ["Tooltip", "validate_integer", "validate_page_number", "validate_path"]
