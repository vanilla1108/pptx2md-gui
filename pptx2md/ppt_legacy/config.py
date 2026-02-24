"""PPT Legacy 转换配置模型。"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Literal, Optional


TableHeaderMode = Literal["first-row", "empty"]


class ConversionCancelled(Exception):
    """用户取消转换。"""


@dataclass(frozen=True)
class ExtractConfig:
    """提取流程配置。"""

    input_path: str
    output_path: Optional[str] = None
    debug: bool = False
    ui: bool = True
    extract_images: bool = True
    image_dir: Optional[str] = None
    table_header: TableHeaderMode = "first-row"

    def __post_init__(self) -> None:
        if str(self.table_header) not in ("first-row", "empty"):
            raise ValueError(f"不支持的 table_header: {self.table_header}")
        if not str(self.input_path or "").strip():
            raise ValueError("input_path 不能为空")

    @classmethod
    def from_cli_args(cls, args) -> "ExtractConfig":
        return cls(
            input_path=args.input,
            output_path=args.output,
            debug=bool(args.debug),
            ui=(not bool(args.no_ui)),
            extract_images=(not bool(args.no_image)),
            image_dir=args.image_dir,
            table_header=args.table_header,
        )
