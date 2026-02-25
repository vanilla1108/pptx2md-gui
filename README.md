# pptx2md-gui

PPT/PPTX 转 Markdown 桌面工具，提供 GUI（仅 Windows）与 CLI，基于 [ssine/pptx2md](https://github.com/ssine/pptx2md) 修改。

> **⚠️ 声明**
>
> 本项目为**个人学习项目**，代码和功能尚不成熟，可能存在较多 Bug。
> 不保证转换效果，也不承诺持续维护与更新。
> 欢迎提 Issue 反馈问题，但修复时间不做保证。如果对你有帮助，欢迎自行 Fork 和修改。

## 功能

- **输入**：`.pptx`、`.ppt`（实验性，需 Windows + PowerPoint）
- **输出**：Markdown（默认）、TiddlyWiki、Madoko、Quarto
- **支持内容**：标题、列表、强调样式、超链接、表格、图片、备注、多栏检测、部分公式与 OLE 预览图
- **GUI 特性**：批量转换、拖放导入、参数预设、深色/浅色主题

## 快速开始

### GUI（Windows 便携版）

从 [Releases](https://github.com/vanilla1108/pptx2md-gui/releases) 下载 zip，解压后运行 exe 即可。

### 源码运行

```bash
pip install -e .

# CLI
python -m pptx2md input.pptx -o out.md

# GUI
python -m pptx2md_gui
```

### 可选依赖

- `wand` + ImageMagick — 提高 WMF 图片转换成功率
- Microsoft PowerPoint — `.ppt` 转换和 WMF COM 兜底必需

## 已知限制

- `.ppt` 为实验性功能，复杂版式可能不稳定
- WMF 转换效果取决于系统环境（Pillow → ImageMagick → PowerPoint COM，命中即停止）
- GUI 仅在 Windows 下验证，其他系统未经测试

## 开发

### 环境搭建

```bash
# 开发
conda activate pptx2md_gui_dev
pip install -e ".[dev]"

# 打包
conda activate gui_build
pip install -e ".[build]"
```

### 测试

```bash
# 快速回归
python -m pytest tests -q -m "not slow and not ppt_com"

# 全量（需要 test_pptx/ 下的样本文件）
python -m pytest tests -q
```

### GUI 打包

```bash
conda activate gui_build
python build_exe.py
```

### 目录结构

| 目录 | 说明 |
|------|------|
| `pptx2md/` | 核心转换库 |
| `pptx2md_gui/` | GUI 应用 |
| `tests/` | 测试 |
| `ppt2md_script/` | 旧版 COM 脚本 |

详细文档：[CLI](CLI_README.md) · [GUI](GUI_README.md) · [测试](tests/README.md)

## 致谢

- [ssine/pptx2md](https://github.com/ssine/pptx2md)
- [dwml](https://pypi.org/project/dwml/)

## 许可证

Apache License 2.0，见 [LICENSE](LICENSE)。
