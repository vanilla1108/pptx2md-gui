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

import os
import re
import urllib.parse
from typing import List

from rapidfuzz import fuzz

from pptx2md.types import ConversionConfig, ElementType, ListType, ParagraphElement, ParsedPresentation, SlideElement, SlideType, TextRun
from pptx2md.utils import rgb_to_hex


class Formatter:

    def __init__(self, config: ConversionConfig):
        os.makedirs(config.output_path.parent, exist_ok=True)
        self.ofile = open(config.output_path, 'w', encoding='utf8', newline='')
        self.config = config

    def output(self, presentation_data: ParsedPresentation):
        self.put_header()

        last_element = None
        last_title = None
        ordered_counters = {}
        first_title_seen = False
        for slide_idx, slide in enumerate(presentation_data.slides):
            # 输出幻灯片编号注释（1-based）
            if not self.config.disable_slide_number:
                self.write(f'<!-- slide: {slide_idx + 1} -->\n')

            all_elements = []
            if slide.type == SlideType.General:
                all_elements = slide.elements
            elif slide.type == SlideType.MultiColumn:
                all_elements = slide.preface + slide.columns

            for elem_idx, element in enumerate(all_elements):
                next_elem = all_elements[elem_idx + 1] if elem_idx + 1 < len(all_elements) else None

                if last_element and last_element.type == ElementType.ListItem and element.type != ElementType.ListItem:
                    self.put_list_footer()
                    ordered_counters = {}

                match element.type:
                    case ElementType.Title:
                        element.content = element.content.strip()
                        if element.content:
                            # 首个标题保持原级别，后续标题降一级
                            effective_level = element.level if not first_title_seen else element.level + 1
                            if last_title and last_title.level == element.level and fuzz.ratio(
                                    last_title.content, element.content, score_cutoff=92):
                                # 如果标题与上一个相同则跳过
                                # 允许重复幻灯片标题 - 添加 (cont.) 后缀
                                if self.config.keep_similar_titles:
                                    self.put_title(f'{element.content} (cont.)', effective_level)
                            else:
                                self.put_title(element.content, effective_level)
                            last_title = element
                            first_title_seen = True
                    case ElementType.ListItem:
                        if not (last_element and last_element.type == ElementType.ListItem):
                            # 前一个元素是段落时跳过列表头空行（段落已提供换行）
                            if not (last_element and last_element.type == ElementType.Paragraph):
                                self.put_list_header()
                            ordered_counters = {}
                        list_type = getattr(element, 'list_type', ListType.Unordered)
                        if list_type == ListType.Ordered:
                            level = element.level
                            explicit_number = getattr(element, 'list_number', None)
                            ordered_counters[level] = self._resolve_ordered_list_number(
                                ordered_counters, level, explicit_number
                            )
                            self.put_list(self.get_formatted_runs(element.content), element.level,
                                          list_type=list_type, list_number=ordered_counters[level])
                        else:
                            self.put_list(self.get_formatted_runs(element.content), element.level)
                    case ElementType.Paragraph:
                        text = self.get_formatted_runs(element.content)
                        # 段落后紧跟列表时，用单换行实现紧凑排版
                        if next_elem and next_elem.type == ElementType.ListItem:
                            self.write(text + '\n')
                        else:
                            self.put_para(text)
                    case ElementType.Image:
                        self.put_image(element.path, element.width)
                    case ElementType.Table:
                        self.put_table([[self.get_formatted_runs(cell) for cell in row] for row in element.content])
                last_element = element

            notes_written = not self.config.disable_notes and slide.notes
            if notes_written:
                self.put_para('---')
                for note in slide.notes:
                    self.put_para(note)

            if slide_idx < len(presentation_data.slides) - 1 and self.config.enable_slides:
                # 列表项以 \n 结尾，补一个 \n 保证分隔符前有一行空行
                if last_element and last_element.type == ElementType.ListItem and not notes_written:
                    self.write('\n')
                    ordered_counters = {}
                self.write('---\n\n')
                # 重置为虚拟段落，防止跨幻灯片列表头/尾产生多余空行
                last_element = ParagraphElement(content=[])

        self.close()

    def put_header(self):
        pass

    @staticmethod
    def _resolve_ordered_list_number(ordered_counters, level, explicit_number=None):
        current = ordered_counters.get(level)
        if explicit_number is None:
            ordered_counters[level] = (current or 0) + 1
        elif current is None:
            # 首项使用显式 startAt 作为种子。
            ordered_counters[level] = explicit_number
        elif explicit_number <= current:
            # 部分 PPT 模板会给同一连续列表的每一项重复写入相同 startAt，
            # 这里按“连续编号”推进，避免 3,3,3,3。
            ordered_counters[level] = current + 1
        else:
            # 显式编号向前跳跃（如手动重设为更大数字）时，尊重源文档。
            ordered_counters[level] = explicit_number

        for deeper_level in [k for k in ordered_counters if k > level]:
            ordered_counters.pop(deeper_level, None)

        return ordered_counters[level]

    def put_title(self, text, level):
        pass

    def put_list(self, text, level, list_type=ListType.Unordered, list_number=None):
        pass

    def put_list_header(self):
        self.put_para('')

    def put_list_footer(self):
        self.put_para('')

    def get_formatted_runs(self, runs: List[TextRun]):
        res = ''
        for run in runs:
            text = run.text
            if text == '':
                continue

            if run.style.is_math:
                res += self.get_math(text)
                continue

            if not self.config.disable_escaping:
                text = self.get_escaped(text)

            if run.style.hyperlink:
                text = self.get_hyperlink(text, run.style.hyperlink)
            if run.style.is_accent:
                text = self.get_accent(text)
            elif run.style.is_strong:
                text = self.get_strong(text)
            if run.style.color_rgb and not self.config.disable_color:
                text = self.get_colored(text, run.style.color_rgb)

            res += text
        return res.strip()

    def put_para(self, text):
        pass

    def put_image(self, path, max_width):
        pass

    def put_table(self, table):
        pass

    def get_accent(self, text):
        pass

    def get_strong(self, text):
        pass

    def get_colored(self, text, rgb):
        pass

    def get_hyperlink(self, text, url):
        pass

    def get_escaped(self, text):
        pass

    def get_math(self, text):
        return f' ${text}$ '

    def write(self, text):
        self.ofile.write(text)

    def flush(self):
        self.ofile.flush()

    def close(self):
        self.ofile.close()
        if self.config.compress_blank_lines:
            _compress_output_blank_lines(self.config.output_path)


def _compress_output_blank_lines(output_path):
    """将输出文件中的连续空行压缩为 1 行空行。"""
    with open(output_path, 'r', encoding='utf8', newline='') as f:
        original = f.read()

    # 统一换行符后进行处理，避免 Windows \r\n 与 \n 混用导致边界错误。
    normalized = original.replace('\r\n', '\n')
    lines = normalized.split('\n')

    out_lines = []
    last_was_blank = False
    for line in lines:
        is_blank = line.strip() == ''
        if is_blank:
            if last_was_blank:
                continue
            out_lines.append('')
            last_was_blank = True
        else:
            out_lines.append(line)
            last_was_blank = False

    compressed = '\n'.join(out_lines)

    # 保持“文件末尾是否有换行”与原始一致。
    if original.endswith('\n') or original.endswith('\r\n'):
        if not compressed.endswith('\n'):
            compressed += '\n'

    if compressed == normalized:
        return

    with open(output_path, 'w', encoding='utf8', newline='') as f:
        f.write(compressed)


class MarkdownFormatter(Formatter):
    # 将输出写入 Markdown 格式
    def __init__(self, config: ConversionConfig):
        super().__init__(config)
        self.esc_re1 = re.compile(r'([\\\*`!_\{\}\[\]\(\)#\+-\.])')
        self.esc_re2 = re.compile(r'(<[^>]+>)')

    def put_title(self, text, level):
        self.ofile.write('#' * level + ' ' + text + '\n\n')

    def put_list(self, text, level, list_type=ListType.Unordered, list_number=None):
        indent = '  ' * level
        if list_type == ListType.Ordered and list_number is not None:
            self.ofile.write(f'{indent}{list_number}. {text.strip()}\n')
        else:
            self.ofile.write(indent + '* ' + text.strip() + '\n')

    def put_para(self, text):
        self.ofile.write(text + '\n\n')

    def put_image(self, path, max_width=None):
        if max_width is None:
            self.ofile.write(f'![]({urllib.parse.quote(path)})\n\n')
        else:
            self.ofile.write(f'<img src="{path}" style="max-width:{max_width}px;" />\n\n')

    def put_table(self, table):
        gen_table_row = lambda row: '| ' + ' | '.join([c.replace('\n', '<br />') for c in row]) + ' |'
        self.ofile.write(gen_table_row(table[0]) + '\n')
        self.ofile.write(gen_table_row([':-:' for _ in table[0]]) + '\n')
        self.ofile.write('\n'.join([gen_table_row(row) for row in table[1:]]) + '\n\n')

    def get_accent(self, text):
        return ' _' + text + '_ '

    def get_strong(self, text):
        return ' __' + text + '__ '

    def get_colored(self, text, rgb):
        return ' <span style="color:%s">%s</span> ' % (rgb_to_hex(rgb), text)

    def get_hyperlink(self, text, url):
        return '[' + text + '](' + url + ')'

    def esc_repl(self, match):
        return '\\' + match.group(0)

    def get_escaped(self, text):
        text = re.sub(self.esc_re1, self.esc_repl, text)
        text = re.sub(self.esc_re2, self.esc_repl, text)
        return text


class WikiFormatter(Formatter):
    # 将输出写入 wikitext 格式
    def __init__(self, config: ConversionConfig):
        super().__init__(config)
        self.esc_re = re.compile(r'<([^>]+)>')

    def put_title(self, text, level):
        self.ofile.write('!' * level + ' ' + text + '\n\n')

    def put_list(self, text, level, list_type=ListType.Unordered, list_number=None):
        if list_type == ListType.Ordered:
            self.ofile.write('#' * (level + 1) + ' ' + text.strip() + '\n')
        else:
            self.ofile.write('*' * (level + 1) + ' ' + text.strip() + '\n')

    def put_para(self, text):
        self.ofile.write(text + '\n\n')

    def put_image(self, path, max_width):
        if max_width is None:
            self.ofile.write(f'<img src="{path}" />\n\n')
        else:
            self.ofile.write(f'<img src="{path}" width={max_width}px />\n\n')

    def get_accent(self, text):
        return ' __' + text + '__ '

    def get_strong(self, text):
        return ' \'\'' + text + '\'\' '

    def get_colored(self, text, rgb):
        return ' @@color:%s; %s @@ ' % (rgb_to_hex(rgb), text)

    def get_hyperlink(self, text, url):
        return '[[' + text + '|' + url + ']]'

    def esc_repl(self, match):
        return "''''" + match.group(0)

    def get_escaped(self, text):
        text = re.sub(self.esc_re, self.esc_repl, text)
        return text


class MadokoFormatter(Formatter):
    # 将输出写入 Madoko Markdown 格式
    def __init__(self, config: ConversionConfig):
        super().__init__(config)
        self.ofile.write('[TOC]\n\n')
        self.esc_re1 = re.compile(r'([\\\*`!_\{\}\[\]\(\)#\+-\.])')
        self.esc_re2 = re.compile(r'(<[^>]+>)')

    def put_title(self, text, level):
        self.ofile.write('#' * level + ' ' + text + '\n\n')

    def put_list(self, text, level, list_type=ListType.Unordered, list_number=None):
        indent = '  ' * level
        if list_type == ListType.Ordered and list_number is not None:
            self.ofile.write(f'{indent}{list_number}. {text.strip()}\n')
        else:
            self.ofile.write(indent + '* ' + text.strip() + '\n')

    def put_para(self, text):
        self.ofile.write(text + '\n\n')

    def put_image(self, path, max_width):
        if max_width is None:
            self.ofile.write(f'<img src="{path}" />\n\n')
        elif max_width < 500:
            self.ofile.write(f'<img src="{path}" width={max_width}px />\n\n')
        else:
            self.ofile.write('~ Figure {caption: image caption}\n')
            self.ofile.write('![](%s){width:%spx;}\n' % (path, max_width))
            self.ofile.write('~\n\n')

    def get_accent(self, text):
        return ' _' + text + '_ '

    def get_strong(self, text):
        return ' __' + text + '__ '

    def get_colored(self, text, rgb):
        return ' <span style="color:%s">%s</span> ' % (rgb_to_hex(rgb), text)

    def get_hyperlink(self, text, url):
        return '[' + text + '](' + url + ')'

    def esc_repl(self, match):
        return '\\' + match.group(0)

    def get_escaped(self, text):
        text = re.sub(self.esc_re1, self.esc_repl, text)
        text = re.sub(self.esc_re2, self.esc_repl, text)
        return text


class QuartoFormatter(Formatter):
    # 将输出写入 Quarto Markdown 格式 - reveal.js
    def __init__(self, config: ConversionConfig):
        super().__init__(config)
        self.esc_re1 = re.compile(r'([\\\*`!_\{\}\[\]\(\)#\+-\.])')
        self.esc_re2 = re.compile(r'(<[^>]+>)')

    def output(self, presentation_data: ParsedPresentation):
        self.put_header()

        last_title = None
        first_title_seen = False

        def put_elements(elements: List[SlideElement], initial_last_element=None):
            nonlocal last_title, first_title_seen

            last_element = initial_last_element
            ordered_counters = {}
            for elem_idx, element in enumerate(elements):
                next_elem = elements[elem_idx + 1] if elem_idx + 1 < len(elements) else None

                if last_element and last_element.type == ElementType.ListItem and element.type != ElementType.ListItem:
                    self.put_list_footer()
                    ordered_counters = {}

                match element.type:
                    case ElementType.Title:
                        element.content = element.content.strip()
                        if element.content:
                            # 首个标题保持原级别，后续标题降一级
                            effective_level = element.level if not first_title_seen else element.level + 1
                            if last_title and last_title.level == element.level and fuzz.ratio(
                                    last_title.content, element.content, score_cutoff=92):
                                # 如果标题与上一个相同则跳过
                                # 允许重复幻灯片标题 - 添加 (cont.) 后缀
                                if self.config.keep_similar_titles:
                                    self.put_title(f'{element.content} (cont.)', effective_level)
                            else:
                                self.put_title(element.content, effective_level)
                            last_title = element
                            first_title_seen = True
                    case ElementType.ListItem:
                        if not (last_element and last_element.type == ElementType.ListItem):
                            if not (last_element and last_element.type == ElementType.Paragraph):
                                self.put_list_header()
                            ordered_counters = {}
                        list_type = getattr(element, 'list_type', ListType.Unordered)
                        if list_type == ListType.Ordered:
                            level = element.level
                            explicit_number = getattr(element, 'list_number', None)
                            ordered_counters[level] = self._resolve_ordered_list_number(
                                ordered_counters, level, explicit_number
                            )
                            self.put_list(self.get_formatted_runs(element.content), element.level,
                                          list_type=list_type, list_number=ordered_counters[level])
                        else:
                            self.put_list(self.get_formatted_runs(element.content), element.level)
                    case ElementType.Paragraph:
                        text = self.get_formatted_runs(element.content)
                        if next_elem and next_elem.type == ElementType.ListItem:
                            self.write(text + '\n')
                        else:
                            self.put_para(text)
                    case ElementType.Image:
                        self.put_image(element.path, element.width)
                    case ElementType.Table:
                        self.put_table([[self.get_formatted_runs(cell) for cell in row] for row in element.content])
                last_element = element
            return last_element

        after_separator = False
        for slide_idx, slide in enumerate(presentation_data.slides):
            # 输出幻灯片编号注释（1-based）
            if not self.config.disable_slide_number:
                self.write(f'<!-- slide: {slide_idx + 1} -->\n')

            init_last = ParagraphElement(content=[]) if after_separator else None
            slide_last_element = None

            if slide.type == SlideType.General:
                slide_last_element = put_elements(slide.elements, init_last)
            elif slide.type == SlideType.MultiColumn:
                put_elements(slide.preface, init_last)
                if len(slide.columns) == 2:
                    width = '50%'
                elif len(slide.columns) == 3:
                    width = '33%'
                else:
                    raise ValueError(f'Unsupported number of columns: {len(slide.columns)}')

                self.put_para(':::: {.columns}')
                for column in slide.columns:
                    self.put_para(f'::: {{.column width="{width}"}}')
                    put_elements(column)
                    self.put_para(':::')
                self.put_para('::::')

            notes_written = not self.config.disable_notes and slide.notes
            if notes_written:
                self.put_para("::: {.notes}")
                for note in slide.notes:
                    self.put_para(note)
                self.put_para(":::")

            after_separator = False
            if slide_idx < len(presentation_data.slides) - 1 and self.config.enable_slides:
                if slide_last_element and slide_last_element.type == ElementType.ListItem and not notes_written:
                    self.write('\n')
                self.write('---\n\n')
                after_separator = True

        self.close()

    def put_header(self):
        self.ofile.write('''---
title: "Presentation Title"
author: "Author"
format:
  revealjs:
    slide-number: c/t
    width: 1600
    height: 900
    logo: img/logo.png
    footer: "Organization"
    incremental: true
    theme: [simple]
---
''')

    def put_title(self, text, level):
        self.ofile.write('#' * level + ' ' + text + '\n\n')

    def put_list(self, text, level, list_type=ListType.Unordered, list_number=None):
        indent = '  ' * level
        if list_type == ListType.Ordered and list_number is not None:
            self.ofile.write(f'{indent}{list_number}. {text.strip()}\n')
        else:
            self.ofile.write(indent + '* ' + text.strip() + '\n')

    def put_para(self, text):
        self.ofile.write(text + '\n\n')

    def put_image(self, path, max_width=None):
        if max_width is None:
            self.ofile.write(f'![]({urllib.parse.quote(path)})\n\n')
        else:
            self.ofile.write(f'<img src="{path}" style="max-width:{max_width}px;" />\n\n')

    def put_table(self, table):
        gen_table_row = lambda row: '| ' + ' | '.join([c.replace('\n', '<br />') for c in row]) + ' |'
        self.ofile.write(gen_table_row(table[0]) + '\n')
        self.ofile.write(gen_table_row([':-:' for _ in table[0]]) + '\n')
        self.ofile.write('\n'.join([gen_table_row(row) for row in table[1:]]) + '\n\n')

    def get_accent(self, text):
        return ' _' + text + '_ '

    def get_strong(self, text):
        return ' __' + text + '__ '

    def get_colored(self, text, rgb):
        return ' <span style="color:%s">%s</span> ' % (rgb_to_hex(rgb), text)

    def get_hyperlink(self, text, url):
        return '[' + text + '](' + url + ')'

    def esc_repl(self, match):
        return '\\' + match.group(0)

    def get_escaped(self, text):
        text = re.sub(self.esc_re1, self.esc_repl, text)
        text = re.sub(self.esc_re2, self.esc_repl, text)
        return text
