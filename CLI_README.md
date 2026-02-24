# CLI 使用与参数说明

本文档对应 `pptx2md/__main__.py` 当前实现，优先描述“实际行为”，不是历史版本说明。

## 入口

- 可执行脚本：`pptx2md`
- 模块入口：`python -m pptx2md`

## 基本用法

```bash
# .pptx
pptx2md input.pptx -o out.md

# .ppt（实验性，Windows + PowerPoint）
pptx2md input.ppt -o out.md
```

## 输出格式

- Markdown（默认）
- TiddlyWiki：`--wiki`
- Madoko：`--mdk`
- Quarto：`--qmd`

说明：
- `.ppt` 路由仅支持 Markdown；若传入 `--wiki/--mdk/--qmd` 会直接报错退出。

## 参数总览

### 通用输入输出

| 参数 | 说明 |
|---|---|
| `pptx_path` | 输入文件路径（支持 `.pptx` / `.ppt`） |
| `-o, --output` | 输出文件路径 |
| `-i, --image-dir` | 图片输出目录 |
| `-t, --title` | 自定义标题层级文件 |

### 内容与格式控制（主要用于 `.pptx`）

| 参数 | 说明 |
|---|---|
| `--image-width` | 图片最大宽度（px） |
| `--disable-image` | 禁用图片提取 |
| `--disable-wmf` | 保留 WMF 原格式，不做转换 |
| `--color / --no-color` | 开启/关闭颜色标签 |
| `--escaping / --no-escaping` | 开启/关闭 Markdown 特殊字符转义 |
| `--notes / --no-notes` | 开启/关闭备注提取 |
| `--slides` | 添加幻灯片分隔线 `---` |
| `--slide-number / --no-slide-number` | 开启/关闭 slide 注释编号 |
| `--try-multi-column` | 尝试多栏检测（较慢） |
| `--min-block-size` | 文本块最小字符数 |
| `--page` | 仅转换指定页 |
| `--keep-similar-titles` | 保留相似标题并加 `(cont.)` |
| `--no-compress-blank-lines` | 不压缩连续空行 |
| `--wiki` / `--mdk` / `--qmd` | 输出目标格式 |

### `.ppt` 专用参数（COM 路由）

| 参数 | 说明 |
|---|---|
| `--ppt-debug` | 输出 COM 调试日志 |
| `--ppt-no-ui` | PowerPoint 尽量后台运行 |
| `--ppt-table-header {first-row,empty}` | 表格表头策略 |

## 旧参数兼容

为兼容历史命令，以下旧参数仍可使用：

- `--disable-color`
- `--disable-escaping`
- `--disable-notes`
- `--disable-slide-number`
- `--enable-slides`（等价 `--slides`）

## 默认行为差异

### `.pptx` 路由

- 未指定 `-o` 时：
  - 普通 Markdown/Madoko 默认 `out.md`
  - Wiki 默认 `out.tid`
  - Quarto 默认 `out.qmd`
- 未指定 `-i` 时：默认 `输出文件同级/img`

### `.ppt` 路由

- 未指定 `-o` 时：默认输出到当前工作目录，文件名为 `<源文件名>.md`
- 若目标文件已存在：自动追加 `_1/_2/...` 避免覆盖
- 对 `.ppt` 无效的参数（如 `--page`、`--try-multi-column`、`--title`、`--image-width`）会给出警告并忽略

## 示例

```bash
# 1) 仅转换某一页
pptx2md course.pptx --page 10 -o page10.md

# 2) 关闭图片和备注
pptx2md course.pptx --disable-image --no-notes -o text_only.md

# 3) 生成 Quarto
pptx2md course.pptx --qmd -o slides.qmd

# 4) 转 .ppt（实验性）
pptx2md old_deck.ppt --ppt-no-ui --ppt-table-header first-row -o old_deck.md
```

## WMF 与 COM 相关补充

WMF 转换会按以下链路尝试：Pillow -> Wand -> ImageMagick CLI -> PowerPoint COM。

涉及可选环境：
- `pywin32`：PowerPoint COM 相关
- `wand` + ImageMagick：提升 WMF 转换成功率
- Microsoft PowerPoint：`.ppt` 路由与 WMF COM 兜底都需要

可通过环境变量调节 WMF 行为（见 `pptx2md/parser.py`）：
- `PPTX2MD_WMF_COM_FALLBACK`
- `PPTX2MD_WMF_COM_EXPORT_WIDTH`
- `PPTX2MD_WMF_DPI`
- `PPTX2MD_WMF_RASTER_EXT`
- `PPTX2MD_WMF_JPEG_QUALITY`

## API 调用

```python
from pathlib import Path
from pptx2md import convert, ConversionConfig

config = ConversionConfig(
    pptx_path=Path("input.pptx"),
    output_path=Path("out.md"),
    image_dir=Path("img"),
)
convert(config)
```

## 相关文档

- 总览：`README.md`
- GUI：`GUI_README.md`
- 测试：`tests/README.md`
