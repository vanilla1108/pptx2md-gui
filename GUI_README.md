# GUI 使用与开发说明

本文档对应 `pptx2md_gui/` 当前实现。

## 启动

```bash
# 先完成下方“依赖安装（步骤）”
python -m pptx2md_gui
```

## 依赖安装（步骤）

### 开发环境

```bash
conda activate pptx2md_gui_dev
pip install -e ".[dev]"
```

### 打包环境

```bash
conda create -n gui_build python=3.13 pip
conda activate gui_build
pip install -e ".[build]"
```

## GUI 依赖（按实际代码）

### 必需依赖

- `customtkinter`
- `tkinterdnd2`
- `CTkToolTip`
- `Pillow`（日志面板图标绘制）
- `pywin32`（Windows 条件依赖；用于 `.ppt` 转换与 WMF COM 兜底）
- Python 自带 `tkinter`（需包含 Tk）

说明：
- `tkinterdnd2` 在 `pptx2md_gui/app.py` 是顶层导入，未安装会导致 GUI 无法启动。
- 若已安装但 DnD 初始化失败，拖拽会禁用，仍可通过“点击选择文件”使用。

### 运行时可选依赖（由核心转换逻辑触发）

- `wand` + ImageMagick：额外 WMF 转换链路
- Microsoft PowerPoint：`.ppt` 转换必须

可选能力由 `dev` extra 一并提供（含 `wand`）。

## 功能概览

- 支持批量添加 `.pptx/.ppt`
- 参数面板覆盖核心转换参数
- 后台线程转换 + 进度条 + 日志面板
- 预设保存/删除/自动记忆上次配置
- 外观模式（深色/浅色）持久化

## 模块结构

### 入口与主窗口

- `pptx2md_gui/__main__.py`：模块入口
- `pptx2md_gui/app.py`：主窗口、线程调度、日志轮询、预设联动

### 组件层 `components/`

- `file_panel.py`：文件列表与管理
- `drop_zone.py`：拖放/点击导入
- `params_panel.py`：参数编辑面板
- `log_panel.py`：日志显示、进度与开始/取消按钮

### 服务层 `services/`

- `converter.py`：`ConversionWorker` 后台线程
- `config_bridge.py`：GUI 参数与 `ConversionConfig`/`ExtractConfig` 互转
- `preset_manager.py`：预设和外观模式持久化

## 参数映射边界

GUI 中分为两条转换路径：

- `.pptx`：构建 `ConversionConfig`，调用 `pptx2md.convert()`
- `.ppt`：构建 `ppt_legacy.ExtractConfig`，调用 `convert_ppt()`

关键差异：
- `.ppt` 仅支持 Markdown 输出，GUI 里选择其他格式时会自动警告并回退。
- `PPT 转换设置` 分组仅在文件列表中包含 `.ppt` 时启用。

## 预设文件位置

由 `PresetManager` 决定：

- 源码运行：`~/.pptx2md/presets.json`
- 打包后运行（`sys.frozen=True`）：默认 `exe` 同目录的 `presets.json`
- 若 `exe` 目录不可写：自动回退到 `~/.pptx2md/presets.json`

## 打包（Windows）

```powershell
conda activate gui_build
python build_exe.py
```

输出目录：`dist/pptx2md-gui/`

## 开发与测试

```bash
# 测试
python -m pytest -q

# 核心库格式化
make format

# GUI 目录单独格式化
yapf -ir pptx2md_gui/*.py pptx2md_gui/**/*.py
```

说明：当前自动化测试以核心转换逻辑与路由逻辑为主，未覆盖 GUI 交互层端到端 UI 测试。

## 相关文档

- 总览：`README.md`
- CLI：`CLI_README.md`
- 测试：`tests/README.md`
