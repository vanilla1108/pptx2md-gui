"""ppt_legacy.engine 的标准输出流兼容性测试。"""

from pptx2md.ppt_legacy import engine


def test_safe_reconfigure_stream_none():
    """传入 None 时应直接返回，不抛异常。"""
    engine._safe_reconfigure_stream(None)


def test_safe_reconfigure_stream_without_method():
    """流对象不支持 reconfigure 时应静默跳过。"""

    class _NoReconfigure:
        pass

    engine._safe_reconfigure_stream(_NoReconfigure())


def test_safe_reconfigure_stream_calls_method():
    """流对象支持 reconfigure 时应按 UTF-8 参数调用。"""
    called = {}

    class _WithReconfigure:
        def reconfigure(self, **kwargs):
            called.update(kwargs)

    engine._safe_reconfigure_stream(_WithReconfigure())
    assert called == {"encoding": "utf-8", "errors": "replace"}


def test_log_fallback_no_stream_does_not_raise(monkeypatch):
    """无日志回调且 stdout/stderr 为空时，_log 不应崩溃。"""
    monkeypatch.setattr(engine, "_log_cb", None)
    monkeypatch.setattr(engine.sys, "stdout", None)
    monkeypatch.setattr(engine.sys, "stderr", None)

    engine._log("INFO", "hello")
    engine._log("ERROR", "world")

