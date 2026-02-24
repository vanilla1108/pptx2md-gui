"""build_exe.py 行为测试。"""

from types import SimpleNamespace

import build_exe


def test_build_onefile_sets_env_and_reports_exe(monkeypatch, capsys):
    calls = []

    def fake_run(cmd, cwd=None, env=None):
        calls.append({"cmd": cmd, "cwd": cwd, "env": env})
        return SimpleNamespace(returncode=0)

    monkeypatch.setattr(build_exe.subprocess, "run", fake_run)
    monkeypatch.setenv(build_exe.ONEFILE_ENV_VAR, "0")

    ok = build_exe.build(use_onefile=True)
    out = capsys.readouterr().out

    assert ok is True
    assert calls
    assert calls[0]["env"][build_exe.ONEFILE_ENV_VAR] == "1"
    assert str(build_exe.SPEC_FILE) in calls[0]["cmd"]
    assert str(build_exe.DIST_DIR / "pptx2md-gui.exe") in out


def test_build_onedir_forces_non_onefile_env(monkeypatch, capsys):
    calls = []

    def fake_run(cmd, cwd=None, env=None):
        calls.append({"cmd": cmd, "cwd": cwd, "env": env})
        return SimpleNamespace(returncode=0)

    monkeypatch.setattr(build_exe.subprocess, "run", fake_run)
    monkeypatch.setenv(build_exe.ONEFILE_ENV_VAR, "1")

    ok = build_exe.build(use_onefile=False)
    out = capsys.readouterr().out

    assert ok is True
    assert calls
    assert calls[0]["env"][build_exe.ONEFILE_ENV_VAR] == "0"
    assert str(build_exe.DIST_DIR / "pptx2md-gui") in out


def test_build_failure_returns_false(monkeypatch, capsys):
    def fake_run(cmd, cwd=None, env=None):
        return SimpleNamespace(returncode=2)

    monkeypatch.setattr(build_exe.subprocess, "run", fake_run)

    ok = build_exe.build(use_onefile=True)
    out = capsys.readouterr().out

    assert ok is False
    assert "构建失败" in out

