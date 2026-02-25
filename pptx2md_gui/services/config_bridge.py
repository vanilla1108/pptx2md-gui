"""GUI 状态与 ConversionConfig 之间的配置桥接。"""

from pathlib import Path
from typing import TYPE_CHECKING, Any, Dict, Optional

if TYPE_CHECKING:
    from pptx2md.types import ConversionConfig


def _resolve_positive_bool(
    params: Dict[str, Any],
    *,
    positive_key: str,
    legacy_disable_key: str,
    default_when_missing: bool,
) -> bool:
    """优先读取正向键，不存在时回退旧版 disable 键。"""
    if positive_key in params:
        return bool(params[positive_key])
    if legacy_disable_key in params:
        return not bool(params[legacy_disable_key])
    return default_when_missing


def build_config(
    pptx_path: Path,
    params: Dict[str, Any],
    output_dir: Optional[Path] = None,
) -> "ConversionConfig":
    """根据 GUI 参数构建 ConversionConfig。

    参数:
        pptx_path: PPTX 文件路径。
        params: 来自 ParamsPanel.get_params() 的 GUI 参数字典。
        output_dir: 覆盖输出目录（可选）。

    返回:
        可用于转换的 ConversionConfig。
    """
    # 确定输出目录
    if output_dir:
        out_dir = output_dir
    elif params.get("output_dir"):
        out_dir = Path(params["output_dir"])
    else:
        out_dir = pptx_path.parent

    # 确定输出文件名
    stem = pptx_path.stem
    if params.get("naming") == "prefix" and params.get("prefix"):
        stem = f"{params['prefix']}_{stem}"

    # 根据格式确定输出扩展名
    output_format = params.get("output_format", "markdown").lower()
    ext_map = {
        "markdown": ".md",
        "wiki": ".tid",
        "madoko": ".md",
        "quarto": ".qmd",
    }
    ext = ext_map.get(output_format, ".md")

    output_path = out_dir / f"{stem}{ext}"

    # 处理文件名冲突
    output_path = _resolve_conflict(output_path)

    # 确定图片目录
    if params.get("image_dir"):
        image_dir = Path(params["image_dir"])
    else:
        image_dir = out_dir / "img"

    # 解析可选的整数值
    image_width = _parse_int(params.get("image_width"))
    min_block_size = _parse_int(params.get("min_block_size"), default=15)
    page = _parse_int(params.get("page"))

    # 解析标题文件路径
    title_path = None
    if params.get("title_path"):
        title_path = Path(params["title_path"])

    # 确定格式标志
    is_wiki = output_format == "wiki"
    is_mdk = output_format == "madoko"
    is_qmd = output_format == "quarto"
    enable_color = _resolve_positive_bool(
        params,
        positive_key="enable_color",
        legacy_disable_key="disable_color",
        default_when_missing=True,
    )
    enable_escaping = _resolve_positive_bool(
        params,
        positive_key="enable_escaping",
        legacy_disable_key="disable_escaping",
        default_when_missing=True,
    )
    enable_notes = _resolve_positive_bool(
        params,
        positive_key="enable_notes",
        legacy_disable_key="disable_notes",
        default_when_missing=True,
    )
    enable_slide_number = _resolve_positive_bool(
        params,
        positive_key="enable_slide_number",
        legacy_disable_key="disable_slide_number",
        default_when_missing=True,
    )

    from pptx2md.types import ConversionConfig

    return ConversionConfig(
        pptx_path=pptx_path,
        output_path=output_path,
        image_dir=image_dir,
        title_path=title_path,
        image_width=image_width,
        disable_image=params.get("disable_image", False),
        disable_wmf=params.get("disable_wmf", False),
        disable_color=not enable_color,
        disable_escaping=not enable_escaping,
        disable_notes=not enable_notes,
        enable_slides=params.get("enable_slides", False),
        disable_slide_number=not enable_slide_number,
        is_wiki=is_wiki,
        is_mdk=is_mdk,
        is_qmd=is_qmd,
        min_block_size=min_block_size,
        page=page,
        try_multi_column=params.get("try_multi_column", False),
        keep_similar_titles=params.get("keep_similar_titles", False),
        compress_blank_lines=params.get("compress_blank_lines", True),
    )


def load_to_gui(config: "ConversionConfig") -> Dict[str, Any]:
    """将 ConversionConfig 转换为 GUI 参数字典。

    参数:
        config: 要转换的 ConversionConfig。

    返回:
        适用于 ParamsPanel.set_params() 的字典。
    """
    # 从标志确定输出格式
    if config.is_wiki:
        output_format = "wiki"
    elif config.is_mdk:
        output_format = "madoko"
    elif config.is_qmd:
        output_format = "quarto"
    else:
        output_format = "markdown"

    return {
        "output_dir": str(config.output_path.parent) if config.output_path else "",
        "output_format": output_format,
        "disable_image": config.disable_image,
        "disable_wmf": config.disable_wmf,
        "image_dir": str(config.image_dir) if config.image_dir else "",
        "image_width": config.image_width,
        "enable_color": not config.disable_color,
        "enable_escaping": not config.disable_escaping,
        "enable_notes": not config.disable_notes,
        "enable_slides": config.enable_slides,
        "enable_slide_number": not config.disable_slide_number,
        "keep_similar_titles": config.keep_similar_titles,
        "compress_blank_lines": config.compress_blank_lines,
        "min_block_size": config.min_block_size,
        "try_multi_column": config.try_multi_column,
        "page": config.page,
        "title_path": str(config.title_path) if config.title_path else "",
    }


def _parse_int(value: Any, default: Optional[int] = None) -> Optional[int]:
    """将值解析为整数，为空或无效时返回默认值。"""
    if value is None or value == "":
        return default
    try:
        return int(value)
    except (ValueError, TypeError):
        return default


def build_ppt_config(
    ppt_path: Path,
    params: Dict[str, Any],
    output_dir: Optional[Path] = None,
) -> "ExtractConfig":
    """根据 GUI 参数构建 .ppt 转换的 ExtractConfig。

    参数:
        ppt_path: PPT 文件路径。
        params: 来自 ParamsPanel.get_params() 的 GUI 参数字典。
        output_dir: 覆盖输出目录（可选）。

    返回:
        可用于 convert_ppt() 的 ExtractConfig。
    """
    from pptx2md.ppt_legacy.config import ExtractConfig  # 延迟导入

    # 确定输出目录
    if output_dir:
        out_dir = output_dir
    elif params.get("output_dir"):
        out_dir = Path(params["output_dir"])
    else:
        out_dir = ppt_path.parent

    # 确定输出文件名（强制 .md）
    stem = ppt_path.stem
    if params.get("naming") == "prefix" and params.get("prefix"):
        stem = f"{params['prefix']}_{stem}"

    output_path = out_dir / f"{stem}.md"
    output_path = _resolve_conflict(output_path)

    # PPT 专属图片提取开关（独立于 .pptx 的 disable_image）
    ppt_extract_images = params.get("ppt_extract_images", True)

    # PPT 专属图片目录：优先 ppt_image_dir，回退 image_dir，最终回退 out_dir/img
    if params.get("ppt_image_dir"):
        image_dir = str(Path(params["ppt_image_dir"]))
    elif params.get("image_dir"):
        image_dir = str(Path(params["image_dir"]))
    else:
        image_dir = str(out_dir / "img")

    return ExtractConfig(
        input_path=str(ppt_path),
        output_path=str(output_path),
        debug=params.get("ppt_debug", False),
        ui=params.get("ppt_ui", True),
        extract_images=ppt_extract_images,
        image_dir=image_dir,
        table_header=params.get("ppt_table_header", "first-row"),
    )


def _resolve_conflict(path: Path) -> Path:
    """通过追加数字后缀解决文件名冲突。"""
    if not path.exists():
        return path

    stem = path.stem
    suffix = path.suffix
    parent = path.parent

    counter = 1
    while True:
        new_path = parent / f"{stem}_{counter}{suffix}"
        if not new_path.exists():
            return new_path
        counter += 1
