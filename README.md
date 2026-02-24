# pptx2md-gui

`pptx2md` 的增强版仓库：同时提供 `CLI` 和 `GUI`，支持 `.pptx` 与实验性 `.ppt` 转换。

- 核心库目录：`pptx2md/`
- GUI 目录：`pptx2md_gui/`
- 测试目录：`tests/`

## 快速开始

```bash
# 先完成下方“依赖安装（步骤）”
python -m pptx2md input.pptx -o out.md
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

## 依赖说明（按场景）

### 1) 核心转换依赖（CLI/GUI 共用）

以下是 `pyproject.toml` 中直接声明、并被当前代码直接使用的核心依赖：

- `python-pptx`
- `rapidfuzz`
- `Pillow`
- `tqdm`
- `pydantic`
- `dwml`
- `numpy`
- `scipy`
- `pywin32`（Windows 条件依赖）

### 2) GUI 依赖

- `customtkinter`
- `tkinterdnd2`
- `CTkToolTip`
- Python 自带 `tkinter`（需你的 Python 发行版包含 Tk）

说明：
- 当前实现里 `tkinterdnd2` 是导入时硬依赖（`pptx2md_gui/app.py` 顶层导入）。
- 若 `tkinterdnd2` 已安装但运行时 DnD 初始化失败，会退化为“点击选文件”模式。

### 3) 可选依赖（仅特定能力需要）

- `wand` + ImageMagick：提高 WMF 转换成功率（可选链路）。
- Microsoft PowerPoint：`.ppt` 转换和 WMF COM 兜底都需要。

可选能力由 `dev` extra 一并提供（含 `wand`）。

## 功能概览

- 输入格式：`.pptx`、`.ppt`（实验性，Windows + PowerPoint）
- 输出格式：Markdown（默认）、TiddlyWiki（`--wiki`）、Madoko（`--mdk`）、Quarto（`--qmd`）
- 支持内容：标题、列表、强调样式、超链接、表格、图片、备注、多栏检测、部分公式与 OLE 预览图

WMF 转换链路（命中即停止）：

1. Pillow
2. Wand（ImageMagick）
3. `magick` CLI
4. PowerPoint COM 兜底

## 文档导航

- CLI 细节：`CLI_README.md`
- GUI 细节：`GUI_README.md`
- 测试说明：`tests/README.md`
- 旧版 COM 脚本：`ppt2md_script/README.md`

## 测试

```bash
# 快速回归（不依赖慢测样本，不触发 COM 测试）
python -m pytest tests -q -m "not slow and not ppt_com"

# 全量测试（需要 test_pptx 样本文件）
python -m pytest tests -q
```

说明：`slow` 测试依赖 `test_pptx/` 下的样本文件，若样本未放入仓库会报错。

## Windows GUI 打包

```powershell
conda activate gui_build
python build_exe.py
```

输出目录为 `dist/pptx2md-gui/`。

## 已知限制

- `.ppt` 仍是实验性能力，复杂版式存在不稳定性。
- WMF 转换受系统环境影响较大（ImageMagick/COM 可用性）。
- GUI 主要在 Windows 场景验证，其他系统未做持续验证。

## 致谢

- [ssine/pptx2md](https://github.com/ssine/pptx2md)
- [dwml](https://pypi.org/project/dwml/)

## 许可证

Apache License 2.0，见 `LICENSE`。
