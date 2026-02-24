# Copyright 2024 Liu Siyao
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#     http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.
#
# Modifications Copyright 2025-2026 vanilla1108

from __future__ import annotations

from enum import Enum
from pathlib import Path
from typing import List, Optional, Union

from pydantic import BaseModel


class ConversionConfig(BaseModel):
    """PowerPoint 到 Markdown 转换的配置。"""

    pptx_path: Path
    """待转换的 PPTX 文件路径"""

    output_path: Path
    """输出文件路径"""

    image_dir: Optional[Path]
    """提取图片的存放目录"""

    title_path: Optional[Path] = None
    """自定义标题列表文件路径"""

    image_width: Optional[int] = None
    """图片最大宽度（像素）"""

    disable_image: bool = False
    """禁用图片提取"""

    disable_wmf: bool = False
    """保留 WMF 格式图片不转换（避免在 Linux 下出现异常）"""

    disable_color: bool = False
    """不添加颜色 HTML 标签"""

    disable_escaping: bool = False
    """不转义特殊字符"""

    disable_notes: bool = False
    """不添加演讲备注"""

    enable_slides: bool = False
    """使用 `\n---\n` 分隔幻灯片"""

    disable_slide_number: bool = False
    """不在每页幻灯片内容前添加编号注释"""

    is_wiki: bool = False
    """生成 wikitext 格式输出（TiddlyWiki）"""

    is_mdk: bool = False
    """生成 Madoko Markdown 格式输出"""

    is_qmd: bool = False
    """生成 Quarto Markdown 演示文稿格式输出"""

    min_block_size: int = 15
    """文本块被转换的最小字符数"""

    page: Optional[int] = None
    """仅转换指定页码"""

    custom_titles: dict[str, int] = {}
    """自定义标题到标题级别的映射"""

    try_multi_column: bool = False
    """尝试检测多列布局幻灯片"""

    keep_similar_titles: bool = False
    """保留相似标题（允许重复幻灯片标题 - 添加 (cont.) 后缀）"""

    compress_blank_lines: bool = True
    """压缩连续空行（将多行空行合并为 1 行空行）。

    该选项仅影响最终输出文本的格式化，不改变 PPTX 解析/提取逻辑。
    """


class ElementType(str, Enum):
    Title = "Title"
    ListItem = "ListItem"
    Paragraph = "Paragraph"
    Image = "Image"
    Table = "Table"


class ListType(str, Enum):
    Unordered = "Unordered"
    Ordered = "Ordered"


class TextStyle(BaseModel):
    is_accent: bool = False
    is_strong: bool = False
    is_math: bool = False
    color_rgb: Optional[tuple[int, int, int]] = None
    hyperlink: Optional[str] = None


class TextRun(BaseModel):
    text: str
    style: TextStyle


class Position(BaseModel):
    left: float
    top: float
    width: float
    height: float


class BaseElement(BaseModel):
    type: ElementType
    position: Optional[Position] = None
    style: Optional[TextStyle] = None


class TitleElement(BaseElement):
    type: ElementType = ElementType.Title
    content: str
    level: int


class ListItemElement(BaseElement):
    type: ElementType = ElementType.ListItem
    content: List[TextRun]
    level: int = 1
    list_type: ListType = ListType.Unordered
    # 仅用于有序列表：当 PPT 段落显式指定了 buAutoNum@startAt 时，保留该编号。
    list_number: Optional[int] = None


class ParagraphElement(BaseElement):
    type: ElementType = ElementType.Paragraph
    content: List[TextRun]


class ImageElement(BaseElement):
    type: ElementType = ElementType.Image
    path: str
    width: Optional[int] = None
    original_ext: str = ""  # 用于记录原始文件扩展名（如 wmf）
    alt_text: str = ""  # 用于无障碍访问


class TableElement(BaseElement):
    type: ElementType = ElementType.Table
    content: List[List[List[TextRun]]]  # 行 -> 列 -> 富文本


SlideElement = Union[TitleElement, ListItemElement, ParagraphElement, ImageElement, TableElement]


class SlideType(str, Enum):
    MultiColumn = "MultiColumn"
    General = "General"


class MultiColumnSlide(BaseModel):
    type: SlideType = SlideType.MultiColumn
    preface: List[SlideElement]
    columns: List[SlideElement]
    notes: List[str] = []


class GeneralSlide(BaseModel):
    type: SlideType = SlideType.General
    elements: List[SlideElement]
    notes: List[str] = []


Slide = Union[GeneralSlide, MultiColumnSlide]


class ParsedPresentation(BaseModel):
    slides: List[Slide]
