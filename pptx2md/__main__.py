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

import argparse
import logging
import sys
from pathlib import Path

from pptx2md.entry import convert
from pptx2md.log import setup_logging
from pptx2md.types import ConversionConfig

setup_logging(compat_tqdm=True)
logger = logging.getLogger(__name__)


def _resolve_enable_flag(enable_value: bool, legacy_disable_value: bool) -> bool:
    """统一处理新旧参数：旧 disable 参数优先级更高。"""
    return False if legacy_disable_value else enable_value


def parse_args() -> argparse.Namespace:
    """解析命令行参数，返回原始 Namespace（不构建 config）。"""
    arg_parser = argparse.ArgumentParser(description='Convert pptx/ppt to markdown')
    arg_parser.add_argument('pptx_path', type=Path, help='path to the pptx/ppt file to be converted')
    arg_parser.add_argument('-t', '--title', type=Path, help='path to the custom title list file')
    arg_parser.add_argument('-o', '--output', type=Path, help='path of the output file')
    arg_parser.add_argument('-i', '--image-dir', type=Path, help='where to put images extracted')
    arg_parser.add_argument('--image-width', type=int, help='maximum image with in px')
    arg_parser.add_argument('--disable-image', action="store_true", help='disable image extraction')
    arg_parser.add_argument('--disable-wmf',
                            action="store_true",
                            help='keep wmf formatted image untouched(avoid exceptions under linux)')
    arg_parser.add_argument(
        '--color',
        dest='enable_color',
        action=argparse.BooleanOptionalAction,
        default=True,
        help='add color HTML tags (use --no-color to disable)',
    )
    arg_parser.add_argument(
        '--disable-color',
        action="store_true",
        help=argparse.SUPPRESS,
    )
    arg_parser.add_argument(
        '--escaping',
        dest='enable_escaping',
        action=argparse.BooleanOptionalAction,
        default=True,
        help='escape markdown special characters (use --no-escaping to disable)',
    )
    arg_parser.add_argument(
        '--disable-escaping',
        action="store_true",
        help=argparse.SUPPRESS,
    )
    arg_parser.add_argument(
        '--notes',
        dest='enable_notes',
        action=argparse.BooleanOptionalAction,
        default=True,
        help='add presenter notes (use --no-notes to disable)',
    )
    arg_parser.add_argument(
        '--disable-notes',
        action="store_true",
        help=argparse.SUPPRESS,
    )
    arg_parser.add_argument(
        '--slides',
        '--enable-slides',
        dest='enable_slides',
        action="store_true",
        help='deliniate slides `\\n---\\n`',
    )
    arg_parser.add_argument(
        '--slide-number',
        dest='enable_slide_number',
        action=argparse.BooleanOptionalAction,
        default=True,
        help='add slide number comments before each slide (use --no-slide-number to disable)',
    )
    arg_parser.add_argument(
        '--disable-slide-number',
        action="store_true",
        help=argparse.SUPPRESS,
    )
    arg_parser.add_argument('--try-multi-column', action="store_true", help='try to detect multi-column slides')
    arg_parser.add_argument('--wiki', action="store_true", help='generate output as wikitext(TiddlyWiki)')
    arg_parser.add_argument('--mdk', action="store_true", help='generate output as madoko markdown')
    arg_parser.add_argument('--qmd', action="store_true", help='generate output as quarto markdown presentation')
    arg_parser.add_argument('--min-block-size',
                            type=int,
                            default=15,
                            help='the minimum character number of a text block to be converted')
    arg_parser.add_argument("--page", type=int, default=None, help="only convert the specified page")
    arg_parser.add_argument(
        "--keep-similar-titles",
        action="store_true",
        help="keep similar titles (allow for repeated slide titles - One or more - Add (cont.) to the title)")
    arg_parser.add_argument(
        "--no-compress-blank-lines",
        dest="compress_blank_lines",
        action="store_false",
        default=True,
        help="do not compress consecutive blank lines in output",
    )

    # PPT Legacy 专用参数
    arg_parser.add_argument('--ppt-debug', action="store_true", help='[PPT] 输出 COM 调试日志')
    arg_parser.add_argument('--ppt-no-ui', action="store_true", help='[PPT] PowerPoint 后台运行')
    arg_parser.add_argument('--ppt-table-header', choices=['first-row', 'empty'], default='first-row',
                            help='[PPT] 表格表头模式')

    return arg_parser.parse_args()


# ---------------------------------------------------------------------------
# 配置构建
# ---------------------------------------------------------------------------


def _build_pptx_config(args: argparse.Namespace) -> ConversionConfig:
    """从 Namespace 构建 .pptx 的 ConversionConfig。"""
    enable_color = _resolve_enable_flag(args.enable_color, args.disable_color)
    enable_escaping = _resolve_enable_flag(args.enable_escaping, args.disable_escaping)
    enable_notes = _resolve_enable_flag(args.enable_notes, args.disable_notes)
    enable_slide_number = _resolve_enable_flag(args.enable_slide_number, args.disable_slide_number)

    if args.output is None:
        extension = '.tid' if args.wiki else '.qmd' if args.qmd else '.md'
        args.output = Path(f'out{extension}')

    return ConversionConfig(
        pptx_path=args.pptx_path,
        output_path=args.output,
        image_dir=args.image_dir or args.output.parent / 'img',
        title_path=args.title,
        image_width=args.image_width,
        disable_image=args.disable_image,
        disable_wmf=args.disable_wmf,
        disable_color=not enable_color,
        disable_escaping=not enable_escaping,
        disable_notes=not enable_notes,
        enable_slides=args.enable_slides,
        disable_slide_number=not enable_slide_number,
        try_multi_column=args.try_multi_column,
        is_wiki=args.wiki,
        is_mdk=args.mdk,
        is_qmd=args.qmd,
        min_block_size=args.min_block_size,
        page=args.page,
        keep_similar_titles=args.keep_similar_titles,
        compress_blank_lines=args.compress_blank_lines,
    )


def _build_ppt_config(args: argparse.Namespace):
    """从 Namespace 构建 .ppt 的 ExtractConfig。"""
    from pptx2md.ppt_legacy.config import ExtractConfig

    output_path = args.output
    if output_path is None:
        # 约定：.ppt 在未指定 -o 时输出到当前工作目录
        output_path = _next_unique_path(Path.cwd() / f"{args.pptx_path.stem}.md")

    image_dir = str(args.image_dir) if args.image_dir else None

    return ExtractConfig(
        input_path=str(args.pptx_path.resolve()),
        output_path=str(output_path),
        debug=args.ppt_debug,
        ui=not args.ppt_no_ui,
        extract_images=not args.disable_image,
        image_dir=image_dir,
        table_header=args.ppt_table_header,
    )


def _next_unique_path(path: Path) -> Path:
    """若路径存在则自动追加 _1/_2...，避免覆盖。"""
    if not path.exists():
        return path

    stem = path.stem
    suffix = path.suffix
    parent = path.parent
    counter = 1
    while True:
        candidate = parent / f"{stem}_{counter}{suffix}"
        if not candidate.exists():
            return candidate
        counter += 1


# ---------------------------------------------------------------------------
# .ppt 路由辅助
# ---------------------------------------------------------------------------


def _check_ppt_format_conflict(wiki: bool, mdk: bool, qmd: bool):
    """检查 .ppt 不支持的输出格式，冲突时退出。"""
    conflicts = []
    if wiki:
        conflicts.append("--wiki")
    if mdk:
        conflicts.append("--mdk")
    if qmd:
        conflicts.append("--qmd")
    if conflicts:
        print(f"错误：PPT 转换仅支持 Markdown 输出，不兼容 {', '.join(conflicts)}", file=sys.stderr)
        sys.exit(1)


def _warn_unsupported_ppt_params(args: argparse.Namespace):
    """对 .ppt 不支持的参数输出警告到 stderr。"""
    warnings = []
    if args.page is not None:
        warnings.append("--page")
    if args.try_multi_column:
        warnings.append("--try-multi-column")
    if args.title is not None:
        warnings.append("--title")
    if args.image_width is not None:
        warnings.append("--image-width")
    if warnings:
        print(f"警告：以下参数对 .ppt 转换无效，将被忽略：{', '.join(warnings)}", file=sys.stderr)


# ---------------------------------------------------------------------------
# 入口
# ---------------------------------------------------------------------------


def main():
    args = parse_args()
    ext = args.pptx_path.suffix.lower()

    if ext == ".ppt":
        # PPT Legacy COM 管道
        _check_ppt_format_conflict(wiki=args.wiki, mdk=args.mdk, qmd=args.qmd)
        _warn_unsupported_ppt_params(args)

        from pptx2md.ppt_legacy import check_environment, convert_ppt

        ok, reason = check_environment(strict=True)
        if not ok:
            print(f"错误：{reason}", file=sys.stderr)
            sys.exit(2)

        config = _build_ppt_config(args)
        success = convert_ppt(config)
        sys.exit(0 if success else 1)
    elif ext == ".pptx":
        # PPTX 原管道
        config = _build_pptx_config(args)
        convert(config)
    else:
        print(f"错误：不支持的文件格式 '{ext}'，仅支持 .pptx / .ppt", file=sys.stderr)
        sys.exit(2)


if __name__ == '__main__':
    main()
