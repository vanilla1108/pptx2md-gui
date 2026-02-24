"""PPT Legacy COM 转换子包。

⚠ 线程安全：engine 模块使用模块级回调引用，
  同一进程内同一时刻仅允许一个 convert_ppt() 调用。
  若未来需支持并行转换，需改为实例化或传参模式。
"""


def check_environment(strict: bool = False) -> tuple[bool, str]:
    """检测 COM 环境是否可用。

    参数:
        strict: False=快速检查(platform+import)，True=额外尝试 Dispatch。

    返回:
        (可用, 原因描述)。可用时原因为空串。
    """
    import platform
    if platform.system() != "Windows":
        return False, "需要 Windows 操作系统"
    try:
        import win32com.client  # noqa: F401
    except ImportError:
        return False, "需要安装 pywin32（pip install pywin32）"

    if strict:
        app = None
        try:
            # 使用 DispatchEx 创建独立实例，避免绑定并关闭用户现有 PowerPoint 进程
            app = win32com.client.DispatchEx("PowerPoint.Application")
            app.DisplayAlerts = 0
        except Exception as e:
            return False, f"PowerPoint 启动失败：{e}"
        finally:
            if app is not None:
                try:
                    app.Quit()
                except Exception:
                    pass

    return True, ""


def convert_ppt(config, log_callback=None, progress_callback=None, cancel_event=None):
    """延迟导入核心模块，执行 .ppt 转换。

    参数:
        config: ExtractConfig 实例。
        log_callback: (level: str, message: str) -> None
        progress_callback: (current: int, total: int, slide_name: str) -> None
        cancel_event: threading.Event 或 None。

    返回:
        bool: 转换是否成功。
    """
    from .engine import extract_ppt_content
    return extract_ppt_content(config, log_callback, progress_callback, cancel_event)
