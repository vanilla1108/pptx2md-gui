"""PowerPoint COM 绑定诊断辅助。"""

from __future__ import annotations

import os
import re
from typing import Any

POWERPOINT_APPLICATION_PROGID = "PowerPoint.Application"
_COMMON_POWERPOINT_PATHS = (
    r"C:\Program Files\Microsoft Office\Root\Office16\POWERPNT.EXE",
    r"C:\Program Files (x86)\Microsoft Office\root\Office16\POWERPNT.EXE",
    r"C:\Program Files\Microsoft Office\Office16\POWERPNT.EXE",
    r"C:\Program Files (x86)\Microsoft Office\Office16\POWERPNT.EXE",
)


def _extract_executable_path(command: str | None) -> str:
    """从 LocalServer32 命令行中提取可执行文件路径。"""
    if not command:
        return ""

    text = str(command).strip().strip("{}").strip()
    if not text:
        return ""

    if text.startswith('"'):
        end = text.find('"', 1)
        if end > 1:
            return text[1:end].strip()

    match = re.search(r"(?i)\.exe\b", text)
    if match:
        return text[: match.end()].strip()

    return text.split()[0].strip()


def classify_powerpoint_server(path_or_command: str | None) -> str:
    """粗略判断 COM 服务器属于 Microsoft PowerPoint / WPS / 未知。"""
    if not path_or_command:
        return "unknown"

    text = str(path_or_command).strip().lower()
    if not text:
        return "unknown"

    if any(token in text for token in ("wpsoffice.exe", "\\wpp.exe", "wps office", "kingsoft")):
        return "wps"
    if "powerpnt.exe" in text or "microsoft office" in text:
        return "microsoft"
    return "unknown"


def get_registered_powerpoint_com_info() -> dict[str, str]:
    """读取注册表中的 PowerPoint.Application 绑定信息。"""
    info = {
        "progid": POWERPOINT_APPLICATION_PROGID,
        "clsid": "",
        "server_command": "",
        "server_path": "",
        "backup_server_command": "",
        "backup_server_path": "",
        "vendor": "unknown",
    }

    if os.name != "nt":
        return info

    try:
        import winreg

        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"PowerPoint.Application\CLSID") as key:
            info["clsid"], _ = winreg.QueryValueEx(key, None)

        if info["clsid"]:
            clsid_subkey = f"CLSID\\{info['clsid']}\\LocalServer32"
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, clsid_subkey) as key:
                info["server_command"], _ = winreg.QueryValueEx(key, None)
                try:
                    info["backup_server_command"], _ = winreg.QueryValueEx(key, ".ksobak")
                except OSError:
                    pass
    except OSError:
        return info

    info["server_path"] = _extract_executable_path(info["server_command"])
    info["backup_server_path"] = _extract_executable_path(info["backup_server_command"])
    info["vendor"] = classify_powerpoint_server(info["server_path"] or info["server_command"])
    return info


def get_runtime_powerpoint_com_info(app: Any) -> dict[str, str]:
    """读取已创建 COM Application 对象的运行时信息。"""
    info = {
        "name": "",
        "path": "",
        "version": "",
        "vendor": "unknown",
    }

    try:
        info["name"] = str(app.Name or "").strip()
    except Exception:
        pass

    try:
        info["path"] = str(app.Path or "").strip()
    except Exception:
        pass

    try:
        info["version"] = str(app.Version or "").strip()
    except Exception:
        pass

    info["vendor"] = classify_powerpoint_server(info["path"] or info["name"])
    return info


def format_powerpoint_com_target(info: dict[str, str]) -> str:
    """把诊断信息格式化为适合日志/报错的短文本。"""
    vendor = info.get("vendor", "unknown")
    target = (
        info.get("path")
        or info.get("server_path")
        or info.get("server_command")
        or "<unknown>"
    )

    if vendor == "microsoft":
        label = "Microsoft PowerPoint"
    elif vendor == "wps":
        label = "WPS Presentation"
    else:
        label = "unknown COM server"

    return f"{label} ({target})"


def find_microsoft_powerpoint_path(registered_info: dict[str, str] | None = None) -> str:
    """尽量推断本机 Microsoft PowerPoint 可执行文件路径。"""
    info = registered_info or get_registered_powerpoint_com_info()
    candidates: list[str] = []

    def _append_candidate(path: str | None):
        text = str(path or "").strip()
        if not text or text in candidates:
            return
        candidates.append(text)

    if classify_powerpoint_server(info.get("backup_server_path")) == "microsoft":
        _append_candidate(info.get("backup_server_path"))
    if classify_powerpoint_server(info.get("server_path")) == "microsoft":
        _append_candidate(info.get("server_path"))
    for path in _COMMON_POWERPOINT_PATHS:
        _append_candidate(path)

    for path in candidates:
        if os.path.exists(path):
            return path
    return candidates[0] if candidates else ""


def format_powerpoint_regserver_command(powerpoint_path: str) -> str:
    """生成适合 PowerShell 的 PowerPoint 重新注册命令。"""
    path = str(powerpoint_path or "").strip()
    if not path:
        return ""
    return f'& "{path}" /regserver'


def build_powerpoint_com_repair_message(
    registered_info: dict[str, str] | None = None,
    runtime_info: dict[str, str] | None = None,
) -> str:
    """生成 PowerPoint COM 被错误绑定时的详细修复提示。"""
    info = registered_info or get_registered_powerpoint_com_info()
    lines = [
        "检测到 PowerPoint.Application COM 当前没有绑定到 Microsoft PowerPoint。",
    ]

    if info.get("server_path") or info.get("server_command"):
        lines.append(f"当前注册表目标：{format_powerpoint_com_target(info)}")
    if runtime_info:
        lines.append(f"当前实际启动目标：{format_powerpoint_com_target(runtime_info)}")
    if info.get("backup_server_path"):
        lines.append(f"检测到历史 Microsoft PowerPoint 路径：{info['backup_server_path']}")

    powerpnt_path = find_microsoft_powerpoint_path(info)
    regserver_cmd = format_powerpoint_regserver_command(powerpnt_path)

    lines.append("这通常是 COM 自动化注册被 WPS 接管，不等于必须把 .ppt 的文件打开方式改成 PowerPoint。")
    lines.append("如果你希望继续使用 WPS 打开 .ppt 文件，也可以先按下面步骤修复 COM 自动化。")
    lines.append("建议按下面顺序修复：")
    lines.append("1. 退出所有 WPS 和 PowerPoint 进程。")
    lines.append("2. 打开“控制面板 -> 程序 -> 程序和功能”，右键“Microsoft Office”并选择“更改”，然后执行“快速修复”。")
    if regserver_cmd:
        lines.append(f"3. 快速修复完成后，在 PowerShell 中运行：{regserver_cmd}")
    else:
        lines.append("3. 快速修复完成后，运行 Microsoft PowerPoint 安装目录中的 POWERPNT.EXE /regserver")
    lines.append("4. 重新打开本工具后再试一次。")
    lines.append("如果仍然失败，再考虑执行一次 Office 快速修复后重复上面的 /regserver。")

    return "\n".join(lines)
