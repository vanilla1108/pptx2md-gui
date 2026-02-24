"""兼容入口 - 实际实现已迁移到 pptx2md.ppt_legacy。

保留 CLI 入口以便独立调用：python -m ppt2md_script.extract_ppt <PPT文件>
"""

import sys
import argparse

from pptx2md.ppt_legacy.config import ExtractConfig
from pptx2md.ppt_legacy import convert_ppt, check_environment


def build_arg_parser():
    parser = argparse.ArgumentParser(description='提取PPT内容为Markdown')
    parser.add_argument('input', help='PPT文件路径')
    parser.add_argument('-o', '--output', help='输出路径')
    parser.add_argument('--no-image', action='store_true',
                        help='禁用图片提取，输出 ![图片] 占位标注')
    parser.add_argument('--image-dir',
                        help='图片输出目录（默认：输出Markdown目录下的 img）')
    parser.add_argument('--debug', action='store_true', help='输出调试日志到stderr')
    parser.add_argument('--no-ui', action='store_true',
                        help='无UI模式：PowerPoint后台打开')
    parser.add_argument('--table-header', choices=['first-row', 'empty'], default='first-row',
                        help='表格表头模式')
    return parser


def main():
    ok, reason = check_environment()
    if not ok:
        print(f"错误：{reason}", file=sys.stderr)
        sys.exit(2)

    parser = build_arg_parser()
    args = parser.parse_args()
    config = ExtractConfig.from_cli_args(args)
    success = convert_ppt(config)
    sys.exit(0 if success else 1)


if __name__ == '__main__':
    if len(sys.argv) == 1:
        print("用法: python -m ppt2md_script.extract_ppt <PPT文件> [-o 输出路径]")
    else:
        main()
