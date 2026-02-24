"""用于保存和加载转换配置的预设管理器。"""

import json
from pathlib import Path
import sys
from typing import Any, Dict, List, Optional


_LEGACY_CONFIG_DIR = Path.home() / ".pptx2md"
_LEGACY_CONFIG_FILE = _LEGACY_CONFIG_DIR / "presets.json"


def _get_default_config_dir() -> Path:
    """获取预设的默认存储目录。

    - 打包为 exe（PyInstaller 等）时：默认使用 exe 同目录（方便便携/绿色发布）。
    - 以源码/解释器运行时：保持历史行为，使用 ~/.pptx2md/。
    """

    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return _LEGACY_CONFIG_DIR


class PresetManager:
    """管理转换预设（保存/加载/删除）。"""

    DEFAULT_PRESET = {
        # 输出设置（GUI 专有）
        "output_dir": "",
        "naming": "same",
        "prefix": "",
        "output_format": "markdown",
        "disable_image": False,
        "disable_wmf": False,
        "image_width": None,
        "enable_color": True,
        "enable_escaping": True,
        "enable_notes": True,
        "enable_slides": False,
        "enable_slide_number": True,
        "try_multi_column": False,
        "min_block_size": 15,
        "keep_similar_titles": False,
        "compress_blank_lines": True,
    }

    def __init__(self, config_dir: Optional[Path] = None):
        """初始化预设管理器。

        参数:
            config_dir: 存储预设的目录。
                - 打包为 exe 时默认是 exe 同目录。
                - 以源码/解释器运行时默认是 ~/.pptx2md/
        """
        if config_dir is None:
            config_dir = _get_default_config_dir()
        self.config_dir = config_dir
        self.config_file = config_dir / "presets.json"
        self._ensure_config_dir()
        self._data = self._load_data()

    def _ensure_config_dir(self):
        """确保配置目录存在。"""
        # exe 同目录通常已存在，这里 mkdir(exist_ok=True) 不会改变目录结构；
        # 但若用户显式传入了一个不存在的目录，也可正常创建。
        self.config_dir.mkdir(parents=True, exist_ok=True)

    def _load_data(self) -> Dict[str, Any]:
        """从文件加载预设数据。"""
        if self.config_file.exists():
            try:
                with open(self.config_file, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    if "presets" not in data:
                        data["presets"] = {}
                    return data
            except (json.JSONDecodeError, IOError):
                pass

        # 兼容旧版本：之前预设固定在 ~/.pptx2md/presets.json。
        # 当默认目录切到 exe 同目录后，如果新位置还没有文件，则尝试读取旧位置。
        if self.config_file != _LEGACY_CONFIG_FILE and _LEGACY_CONFIG_FILE.exists():
            try:
                with open(_LEGACY_CONFIG_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    if "presets" not in data:
                        data["presets"] = {}

                    # 尝试把旧数据迁移到新位置（最佳努力，不影响正常启动）。
                    if not self.config_file.exists():
                        try:
                            with open(self.config_file, "w", encoding="utf-8") as out:
                                json.dump(data, out, ensure_ascii=False, indent=2)
                        except OSError:
                            pass
                    return data
            except (json.JSONDecodeError, IOError):
                pass

        return {
            "presets": {"默认配置": self.DEFAULT_PRESET.copy()},
            "last_used": "默认配置",
        }

    def _save_data(self):
        """将预设数据保存到文件。"""
        try:
            with open(self.config_file, "w", encoding="utf-8") as f:
                json.dump(self._data, f, ensure_ascii=False, indent=2)
        except OSError:
            # 若 exe 目录不可写（例如装在 Program Files），则回退到旧目录，保证功能可用。
            self.config_dir = _LEGACY_CONFIG_DIR
            self.config_file = _LEGACY_CONFIG_FILE
            self.config_dir.mkdir(parents=True, exist_ok=True)
            with open(self.config_file, "w", encoding="utf-8") as f:
                json.dump(self._data, f, ensure_ascii=False, indent=2)

    def get_preset_names(self) -> List[str]:
        """获取所有预设名称列表。"""
        return list(self._data["presets"].keys())

    def get_preset(self, name: str) -> Optional[Dict[str, Any]]:
        """按名称获取预设。"""
        preset = self._data["presets"].get(name)
        if preset is None:
            return None
        # 兼容旧预设：缺失字段用默认值补齐
        return {**self.DEFAULT_PRESET, **preset}

    def save_preset(self, name: str, params: Dict[str, Any]):
        """保存指定名称的预设。

        参数:
            name: 预设名称。
            params: 要保存的 GUI 参数（仅保存预设相关的键）。
        """
        preset = {}
        for key in self.DEFAULT_PRESET:
            if key in params:
                preset[key] = params[key]
            else:
                preset[key] = self.DEFAULT_PRESET[key]

        self._data["presets"][name] = preset
        self._save_data()

    def delete_preset(self, name: str) -> bool:
        """按名称删除预设。

        参数:
            name: 要删除的预设名称。

        返回:
            删除成功返回 True，未找到或为唯一预设时返回 False。
        """
        if name not in self._data["presets"]:
            return False
        if len(self._data["presets"]) <= 1:
            return False

        del self._data["presets"][name]

        # 如有需要更新 last_used
        if self._data.get("last_used") == name:
            self._data["last_used"] = list(self._data["presets"].keys())[0]

        self._save_data()
        return True

    def get_last_used(self) -> str:
        """获取上次使用的预设名称。"""
        last = self._data.get("last_used", "默认配置")
        if last not in self._data["presets"]:
            last = list(self._data["presets"].keys())[0] if self._data["presets"] else "默认配置"
        return last

    def set_last_used(self, name: str):
        """设置上次使用的预设名称。"""
        if name in self._data["presets"]:
            self._data["last_used"] = name
            self._save_data()

    def get_default_preset(self) -> Dict[str, Any]:
        """获取默认预设值。"""
        return self.DEFAULT_PRESET.copy()

    def get_appearance_mode(self) -> str:
        """获取保存的外观模式（'dark' 或 'light'）。"""
        mode = self._data.get("appearance_mode", "dark")
        return mode if mode in ("dark", "light") else "dark"

    def set_appearance_mode(self, mode: str) -> None:
        """保存外观模式偏好。"""
        self._data["appearance_mode"] = mode
        self._save_data()
