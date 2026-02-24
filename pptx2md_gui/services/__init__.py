"""pptx2md GUI 的业务逻辑服务。"""

from .config_bridge import build_config, load_to_gui
from .converter import ConversionResults, ConversionWorker
from .preset_manager import PresetManager

__all__ = [
    "build_config",
    "load_to_gui",
    "ConversionWorker",
    "ConversionResults",
    "PresetManager",
]
