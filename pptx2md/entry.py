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

import logging
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from pptx2md.types import ConversionConfig

logger = logging.getLogger(__name__)


def convert(config: "ConversionConfig", progress_callback=None, cancel_event=None, disable_tqdm=False):
    """将 PowerPoint 演示文稿转换为 Markdown。

    参数:
        config: 转换配置。
        progress_callback: 可选的进度更新回调，签名: (current, total, slide_name)。
        cancel_event: 可选的 threading.Event，用于支持取消操作。
        disable_tqdm: 禁用 tqdm 进度条（适用于 GUI）。
    """
    import pptx2md.outputter as outputter
    from pptx2md import parser as _parser
    from pptx2md.utils import load_pptx, prepare_titles

    if config.title_path:
        config.custom_titles = prepare_titles(config.title_path)

    prs = load_pptx(config.pptx_path)

    logger.info("conversion started")

    try:
        ast = _parser.parse(
            config,
            prs,
            progress_callback=progress_callback,
            cancel_event=cancel_event,
            disable_tqdm=disable_tqdm,
        )
    finally:
        # 若启用了 WMF 的 COM 兜底，确保转换结束后关闭 PowerPoint 进程
        try:
            _parser.close_powerpoint_com_session()
        except Exception:
            pass

    if cancel_event and cancel_event.is_set():
        logger.warning("Conversion was cancelled")
        return

    if str(config.output_path).endswith('.json'):
        with open(config.output_path, 'w') as f:
            f.write(ast.model_dump_json(indent=2))
        logger.info(f'presentation data saved to {config.output_path}')
        return

    if config.is_wiki:
        out = outputter.WikiFormatter(config)
    elif config.is_mdk:
        out = outputter.MadokoFormatter(config)
    elif config.is_qmd:
        out = outputter.QuartoFormatter(config)
    else:
        out = outputter.MarkdownFormatter(config)

    out.output(ast)
    logger.info(f'converted document saved to {config.output_path}')
