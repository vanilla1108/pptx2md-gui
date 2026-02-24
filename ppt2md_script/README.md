# ppt2md_script

用 Python + PowerPoint（COM/win32com）把 `.ppt/.pptx` 按“视觉阅读顺序”提取文本/表格并输出 Markdown；支持处理嵌入的 PowerPoint 对象（OLE）。

当前实现会尽量做到：

- 单栏页面：按从上到下、从左到右排序
- 常见多栏页面：自动检测左右两栏，按“左栏 -> 右栏”输出（XY-Cut 回退策略）
- 列表：支持项目符号/多级缩进列表；支持“自动编号”列表输出为 Markdown 有序列表

## 环境要求

- Windows
- 已安装 Microsoft PowerPoint（桌面版）
- Python（建议 3.8+）

## 安装与初始化（PowerShell）

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install pywin32
```

## 命令与用法

### 1) 提取 PPT 为 Markdown（主命令）

```powershell
python extract_ppt.py <PPT文件路径> [-o <输出路径>] [--debug] [--no-ui] [--table-header first-row|empty]
```

参数说明：

- `<PPT文件路径>`：必填，输入 `.ppt` 或 `.pptx`。
- `-o, --output <输出路径>`：可选，指定输出 Markdown 文件路径。
  - 不指定时：默认输出到 PPT 所在目录，文件名与 PPT 一致（若同名冲突会自动追加序号）。
- `--debug`：可选，把调试日志与异常堆栈输出到 `stderr`。
- `--no-ui`：可选，尽量无 UI 模式（优先使用 `WithWindow=False` 打开；失败则回退到 `WithWindow=True`）。
  - 注意：不同 PowerPoint 版本/策略差异较大，不保证完全不弹窗。
  - 注意：若遇到“嵌入 PPT 无法直接读取”的情况，会跳过“UI 激活 + DoVerb”回退（可能导致嵌入内容无法提取）。
- `--table-header first-row|empty`：可选，表格表头策略（默认 `first-row`）。
  - `first-row`：首行作为表头输出（默认）。
  - `empty`：输出空表头分隔行，所有行按数据行输出。

退出码：

- `0`：成功
- `1`：失败

### 2) 查看内置用法提示

不带参数运行会打印简要用法：

```powershell
python extract_ppt.py
```

## 输出约定

- 幻灯片标题：默认输出二级标题 `##`；标题来自 `slide.Shapes.Title`，若不存在则按“靠上 + 字号大 + 文本短”择优。
- 路径引用：每页标题下会输出引用块，含页码与路径，便于回溯与定位，例如：

  ```text
  ## 第五-0讲 ARM体系结构
  > 幻灯片 2：`路径：S2`
  ```

- 标题文本框多段：若标题所在文本框里还包含“副标题/正文”（同一 shape 的第 2 段及以后），脚本会在标题后补回这些段落，避免内容丢失。
- 嵌入 PPT：用引用块标记并递归展开（含路径），例如：

  ```text
  > ▶ 嵌入PPT #1 `路径：S2/E1`
  ```

## 常用示例（基于仓库内测试文件）

### 提取到默认输出路径

```powershell
python extract_ppt.py "test-ppt\第5章-0  ARM体系结构.pptx"
```

### 指定输出路径

```powershell
python extract_ppt.py "test-ppt\第5章-0  ARM体系结构.pptx" -o "test-output\第5章-0  ARM体系结构.md"
```

### 输出调试日志

```powershell
python extract_ppt.py "test-ppt\第5章-0  ARM体系结构.pptx" --debug
```

### 表格空表头（所有行按数据输出）

```powershell
python extract_ppt.py "test-ppt\第5章-0  ARM体系结构.pptx" --table-header empty
```

### 后台运行（尽量不弹出窗口）

```powershell
python extract_ppt.py "test-ppt\第5章-0  ARM体系结构.pptx" --no-ui
```

## 已知限制

- 公式/复杂数学排版：PowerPoint 的“公式对象（OMML）”在 COM 的 `TextRange.Text` 中通常会被扁平化为普通字符，复杂结构（括号/分式/上下标布局等）可能无法完整还原。
- 视觉排序：XY-Cut 主要覆盖常见“标题 + 左右两栏正文”等布局；极端/不规则版式仍可能出现顺序不理想的情况。

## 冒烟测试（人工比对）

仓库提供：

- `test-ppt/`：示例 `.ppt/.pptx` 输入
- `test-output/`：示例 Markdown 输出（可用于人工抽查排序/表格格式/嵌入对象）

建议流程：

1. 对 `test-ppt/` 的每个文件运行一次 `extract_ppt.py`（可指定输出到一个临时目录，避免覆盖）。
2. 与 `test-output/` 中对应示例做 diff 或抽查关键点（表格、列表、嵌入 PPT、排序）。
