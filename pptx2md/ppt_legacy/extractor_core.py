"""提取核心公共逻辑（与入口解耦）。"""

import os
import re


_BULLET_LIKE_PREFIXES = [
    "►", "▸", "▹", "▷", "▶", "▻",
    "•", "◦", "‣", "∙", "·", "●", "○",
    "◆", "◇", "■", "□",
    "➤", "➢", "➣", "➔", "→",
    "➜", "➙", "➛", "➝", "➞", "➟",
]

_BULLET_LIKE_PREFIX_RE = re.compile(
    r"^(?:%s)[\s\u00a0]*" % "|".join([re.escape(x) for x in _BULLET_LIKE_PREFIXES])
)

_MANUAL_ORDERED_PREFIX_RE = re.compile(r"^\s*(\d+)\s*、\s*(.+?)\s*$")


def strip_bullet_like_prefix_core(text):
    """若文本以常见“项目符号样式字符”开头，则去掉该符号并返回剩余正文；否则返回 None。"""
    if text is None:
        return None
    s = str(text)
    s = s.replace("\r", " ").replace("\n", " ").replace("\x0b", " ").replace("\t", " ")
    s = s.strip()
    if not s:
        return None

    m = _BULLET_LIKE_PREFIX_RE.match(s)
    if not m:
        return None

    rest = s[m.end():].lstrip(" ")
    return rest if rest else None


def split_manual_ordered_prefix_core(text):
    """识别形如“1、内容”的手打编号，返回 (序号, 正文)。"""
    if text is None:
        return None
    s = str(text)
    s = s.replace("\r", " ").replace("\n", " ").replace("\x0b", " ").replace("\t", " ")
    s = s.strip()
    if not s:
        return None

    m = _MANUAL_ORDERED_PREFIX_RE.match(s)
    if not m:
        return None

    n = int(m.group(1))
    body = m.group(2).strip()
    if not body:
        return None
    return n, body


def looks_like_brief_list_item_core(text, max_len=48):
    """粗略判断该文本是否更像列表项而非完整段落。"""
    if text is None:
        return False
    s = str(text).strip()
    if not s:
        return False
    if len(s) > int(max_len):
        return False
    if re.search(r"[。！？!?]$", s):
        return False
    return True


def escape_md_text_line_core(text):
    """转义会破坏 Markdown 结构的常见情况（用于普通段落行）。"""
    if text is None:
        return ""
    text = str(text)
    text = text.replace("\r", " ")
    text = text.replace("\n", " ")
    text = text.replace("\x0b", " ")
    text = text.replace("\t", " ")

    stripped = text.lstrip(" ")
    prefix = text[:len(text) - len(stripped)]
    if stripped == "":
        return text

    if stripped.startswith(("#", ">")):
        stripped = "\\" + stripped
    elif stripped.startswith(("-", "*", "+")) and len(stripped) >= 2 and stripped[1] == " ":
        stripped = "\\" + stripped
    elif re.match(r"^\d+\.\s", stripped):
        stripped = "\\" + stripped
    elif re.match(r"^(-{3,}|\*{3,}|_{3,})$", stripped.strip()):
        stripped = "\\" + stripped

    return prefix + stripped


def escape_md_table_cell_core(text):
    """转义表格单元格内容，避免破坏管道表格结构。"""
    if text is None:
        return ""
    text = str(text)
    text = text.replace("\r", "<br>")
    text = text.replace("\t", " ")
    text = text.replace("\n", "<br>")
    text = text.replace("\x0b", "<br>")
    text = text.replace("|", "\\|")
    return text.strip()


def get_unique_output_path_core(base_path, path_exists_fn=None):
    """获取唯一的输出路径，如果文件已存在则添加序号。"""
    path_exists_fn = path_exists_fn or os.path.exists
    if not path_exists_fn(base_path):
        return base_path

    dir_name = os.path.dirname(base_path)
    base_name = os.path.splitext(os.path.basename(base_path))[0]
    ext = os.path.splitext(base_path)[1]

    counter = 1
    while True:
        new_path = os.path.join(dir_name, f"{base_name}_{counter}{ext}")
        if not path_exists_fn(new_path):
            return new_path
        counter += 1


def read_shape_alt_text_core(shape, debug_exc_fn=None):
    """读取图片替代文本（失败时返回空串）。"""
    debug_exc_fn = debug_exc_fn or (lambda *_args, **_kwargs: None)
    alt_text = ""
    try:
        alt_text = shape.AlternativeText
    except Exception as e:
        debug_exc_fn("read_shape_alt_text: 读取图片AlternativeText失败", e)
    alt_text = str(alt_text or "").replace("\r", "").replace("\n", " ").strip()
    return alt_text


def safe_shape_id_core(shape):
    """安全读取 shape.Id，失败返回 None。"""
    try:
        return int(shape.Id)
    except Exception:
        return None


def first_paragraph_text_core(shape, debug_exc_fn=None):
    """读取 shape 第一段文本（纯文本，strip 后返回）。"""
    debug_exc_fn = debug_exc_fn or (lambda *_args, **_kwargs: None)
    try:
        if not shape.HasTextFrame or not shape.TextFrame.HasText:
            return None
        tr = shape.TextFrame.TextRange
        if tr.Paragraphs().Count <= 0:
            return None
        return tr.Paragraphs(1, 1).Text.strip() or None
    except Exception as e:
        debug_exc_fn("first_paragraph_text: 读取失败", e)
        return None


def is_title_candidate_shape_core(shape, is_list_block_fn=None, debug_exc_fn=None):
    """判断 shape 是否可能是“标题行”候选（不保证一定是标题）。"""
    debug_exc_fn = debug_exc_fn or (lambda *_args, **_kwargs: None)
    if is_list_block_fn is None:
        raise ValueError("is_list_block_fn 不能为空")

    try:
        if shape.Type in (13, 11):  # 图片
            return False
    except Exception:
        pass

    try:
        if shape.HasTable:
            return False
    except Exception:
        pass

    try:
        if not shape.HasTextFrame or not shape.TextFrame.HasText:
            return False

        if is_list_block_fn(shape):
            return False

        tr = shape.TextFrame.TextRange
        if tr.Paragraphs().Count <= 0:
            return False

        para1 = tr.Paragraphs(1, 1)
        text = para1.Text.strip()
        if not text:
            return False

        try:
            if bool(para1.ParagraphFormat.Bullet.Visible):
                return False
        except Exception:
            pass

        if re.fullmatch(r"\d+", text):
            return False
        if len(text) > 120:
            return False
        return True
    except Exception as e:
        debug_exc_fn("is_title_candidate_shape: 检测失败", e)
        return False


def detect_slide_title_core(
    slide,
    safe_shape_id_fn=None,
    first_paragraph_text_fn=None,
    is_title_candidate_shape_fn=None,
):
    """识别单页幻灯片的第一行标题。"""
    safe_shape_id_fn = safe_shape_id_fn or safe_shape_id_core
    if first_paragraph_text_fn is None:
        raise ValueError("first_paragraph_text_fn 不能为空")
    if is_title_candidate_shape_fn is None:
        raise ValueError("is_title_candidate_shape_fn 不能为空")

    try:
        title_shape = slide.Shapes.Title
        title_text = first_paragraph_text_fn(title_shape)
        if title_text:
            return {"title": title_text, "shape_id": safe_shape_id_fn(title_shape)}
    except Exception:
        pass

    best = None
    try:
        shapes = list(slide.Shapes)
    except Exception:
        shapes = []

    for shape in shapes:
        if not is_title_candidate_shape_fn(shape):
            continue

        text = first_paragraph_text_fn(shape)
        if not text:
            continue

        try:
            top = float(shape.Top)
        except Exception:
            top = 1e9

        try:
            size = float(shape.TextFrame.TextRange.Paragraphs(1, 1).Font.Size)
        except Exception:
            size = 0.0

        score = size * 10.0 - top / 5.0 - len(text) * 0.5
        if top <= 120:
            score += 15.0

        if best is None or score > best["score"]:
            best = {"score": score, "text": text, "shape_id": safe_shape_id_fn(shape)}

    if best is not None:
        return {"title": best["text"], "shape_id": best["shape_id"]}

    return {"title": None, "shape_id": None}


def find_shape_by_id_in_slide_core(slide, shape_id, safe_shape_id_fn=None, debug_exc_fn=None):
    """在 slide.Shapes 中按 shape.Id 查找 Shape（失败返回 None）。"""
    safe_shape_id_fn = safe_shape_id_fn or safe_shape_id_core
    debug_exc_fn = debug_exc_fn or (lambda *_args, **_kwargs: None)
    if shape_id is None:
        return None
    try:
        target = int(shape_id)
    except Exception:
        return None

    try:
        for shape in list(slide.Shapes):
            if safe_shape_id_fn(shape) == target:
                return shape
    except Exception as e:
        debug_exc_fn("find_shape_by_id_in_slide: 枚举Shapes失败", e)
    return None


def extract_title_shape_extra_lines_core(
    slide,
    title_info,
    find_shape_by_id_fn=None,
    extract_text_from_shape_fn=None,
    debug_exc_fn=None,
):
    """补回“标题 shape 里除第一段标题外的其他段落内容”。"""
    debug_exc_fn = debug_exc_fn or (lambda *_args, **_kwargs: None)
    if find_shape_by_id_fn is None:
        raise ValueError("find_shape_by_id_fn 不能为空")
    if extract_text_from_shape_fn is None:
        raise ValueError("extract_text_from_shape_fn 不能为空")

    try:
        if not title_info:
            return []
        sid = title_info.get("shape_id")
        title = title_info.get("title")
        if not sid or not title:
            return []

        shape = find_shape_by_id_fn(slide, sid)
        if shape is None:
            return []

        lines = extract_text_from_shape_fn(shape, skip_first_para_text=str(title))
        title_norm = str(title).strip()
        out = []
        for line in lines:
            if str(line).strip() == title_norm:
                continue
            out.append(line)
        return out
    except Exception as e:
        debug_exc_fn("extract_title_shape_extra_lines: 提取失败", e)
        return []


def build_title_render_context_core(
    slide,
    fallback_title,
    detect_slide_title_fn=None,
    extract_title_shape_extra_lines_fn=None,
):
    """构建标题渲染上下文（标题文本、skip_map、exclude_ids、extra_lines）。"""
    if detect_slide_title_fn is None:
        raise ValueError("detect_slide_title_fn 不能为空")
    if extract_title_shape_extra_lines_fn is None:
        raise ValueError("extract_title_shape_extra_lines_fn 不能为空")

    title_info = detect_slide_title_fn(slide)
    title_text = title_info.get("title") or str(fallback_title)

    skip_map = {}
    exclude_ids = set()
    if title_info.get("shape_id") and title_info.get("title"):
        shape_id = int(title_info["shape_id"])
        skip_map[shape_id] = str(title_info["title"])
        exclude_ids.add(shape_id)

    extra_lines = list(extract_title_shape_extra_lines_fn(slide, title_info) or [])
    return {
        "title_info": title_info,
        "title_text": title_text,
        "skip_map": skip_map,
        "exclude_ids": exclude_ids,
        "extra_lines": extra_lines,
    }


def normalize_md_link_path_core(path):
    """把 Windows 路径规范化为 Markdown 链接路径。"""
    return str(path).replace("\\", "/")


def build_image_placeholder_markdown_core(shape=None, alt_text=None, read_shape_alt_text_fn=None):
    """生成图片占位标注（禁用导图或导图失败时使用）。"""
    if alt_text is None and shape is not None and read_shape_alt_text_fn is not None:
        alt_text = read_shape_alt_text_fn(shape)
    alt_text = str(alt_text or "").strip()
    if alt_text:
        safe_alt = alt_text.replace("]", "\\]")
        return f"![图片: {safe_alt}]"
    return "![图片]"


def next_export_image_path_core(image_ctx, image_loc=None, shape=None, safe_shape_id_fn=None):
    """按顺序生成唯一图片文件名。"""
    image_ctx["counter"] = int(image_ctx.get("counter", 0)) + 1
    idx = int(image_ctx["counter"])

    loc_part = re.sub(r"[^A-Za-z0-9]+", "_", str(image_loc or "").strip()).strip("_")
    if loc_part:
        loc_part = loc_part[:48]

    sid = safe_shape_id_fn(shape) if safe_shape_id_fn is not None else None
    parts = [f"img_{idx:04d}"]
    if loc_part:
        parts.append(loc_part)
    if sid is not None:
        parts.append(f"s{sid}")

    filename = "_".join(parts) + ".png"
    return os.path.join(str(image_ctx.get("dir") or ""), filename)


def build_image_extract_context_core(output_path, extract_images=True, image_dir=None):
    """构建图片导出上下文。"""
    md_dir = os.path.dirname(os.path.abspath(output_path))
    if not extract_images:
        return {"enabled": False, "dir": None, "md_dir": md_dir, "counter": 0}

    if image_dir:
        if os.path.isabs(image_dir):
            target_dir = os.path.abspath(image_dir)
        else:
            target_dir = os.path.abspath(os.path.join(md_dir, image_dir))
    else:
        target_dir = os.path.join(md_dir, "img")

    return {"enabled": True, "dir": target_dir, "md_dir": md_dir, "counter": 0}


def export_shape_image_markdown_core(
    shape,
    image_ctx=None,
    image_loc=None,
    read_shape_alt_text_fn=None,
    build_image_placeholder_markdown_fn=None,
    debug_exc_fn=None,
    makedirs_fn=None,
    next_export_image_path_fn=None,
    wait_com_fn=None,
    com_open_timeout_sec=15,
    file_exists_fn=None,
    relpath_fn=None,
    normalize_md_link_path_fn=None,
):
    """导出 shape 图片并返回 Markdown 图片语法；失败时回退占位标注。"""
    debug_exc_fn = debug_exc_fn or (lambda *_args, **_kwargs: None)
    makedirs_fn = makedirs_fn or os.makedirs
    file_exists_fn = file_exists_fn or os.path.exists
    relpath_fn = relpath_fn or os.path.relpath
    normalize_md_link_path_fn = normalize_md_link_path_fn or normalize_md_link_path_core

    if read_shape_alt_text_fn is None:
        raise ValueError("read_shape_alt_text_fn 不能为空")
    if build_image_placeholder_markdown_fn is None:
        raise ValueError("build_image_placeholder_markdown_fn 不能为空")
    if next_export_image_path_fn is None:
        raise ValueError("next_export_image_path_fn 不能为空")

    alt_text = read_shape_alt_text_fn(shape)
    placeholder = build_image_placeholder_markdown_fn(alt_text=alt_text)

    if not image_ctx or (not image_ctx.get("enabled")):
        return placeholder

    image_dir = image_ctx.get("dir")
    if not image_dir:
        return placeholder

    try:
        makedirs_fn(image_dir, exist_ok=True)
    except Exception as e:
        debug_exc_fn("export_shape_image_markdown: 创建图片目录失败", e)
        return placeholder

    image_path = next_export_image_path_fn(image_ctx, image_loc=image_loc, shape=shape)

    try:
        if wait_com_fn is not None:
            wait_com_fn(
                lambda: shape.Export(image_path, 2),  # 2 = ppShapeFormatPNG
                com_open_timeout_sec,
                "export_shape_image_markdown: shape.Export失败",
            )
        else:
            shape.Export(image_path, 2)
        if not file_exists_fn(image_path):
            raise FileNotFoundError(f"导出后未找到图片文件: {image_path}")
    except Exception as e:
        debug_exc_fn("export_shape_image_markdown: 导出图片失败", e)
        return placeholder

    try:
        md_dir = str(image_ctx.get("md_dir") or "")
        if md_dir:
            link_path = relpath_fn(image_path, start=md_dir)
        else:
            link_path = image_path
    except Exception:
        link_path = image_path

    link_path = normalize_md_link_path_fn(link_path)

    if alt_text:
        safe_alt = alt_text.replace("]", "\\]")
        alt_label = f"图片: {safe_alt}"
    else:
        alt_label = "图片"
    return f"![{alt_label}]({link_path})"


def is_list_block_core(shape, debug_exc_fn=None):
    """检测是否为列表块（参考 pptx2md 的逻辑）。"""
    debug_exc_fn = debug_exc_fn or (lambda *_args, **_kwargs: None)
    try:
        tr = shape.TextFrame.TextRange
        para_count = tr.Paragraphs().Count
        if para_count == 0:
            return False

        levels = set()
        for i in range(1, para_count + 1):
            para = tr.Paragraphs(i, 1)
            level = para.IndentLevel
            levels.add(level)
            if level > 1 or len(levels) > 1:
                return True
        return False
    except Exception as e:
        debug_exc_fn("is_list_block: 读取段落缩进失败", e)
        return False


def get_single_line_plain_text_core(shape, is_list_block_fn=None, debug_exc_fn=None):
    """尝试从 shape 提取“可合并为一行”的纯文本。"""
    debug_exc_fn = debug_exc_fn or (lambda *_args, **_kwargs: None)
    if is_list_block_fn is None:
        raise ValueError("is_list_block_fn 不能为空")

    try:
        if shape.Type in (13, 11):  # 图片
            return None
    except Exception:
        pass

    try:
        if shape.HasTable:
            return None
    except Exception:
        pass

    try:
        if not shape.HasTextFrame:
            return None
        if not shape.TextFrame.HasText:
            return None

        tr = shape.TextFrame.TextRange
        para_count = tr.Paragraphs().Count
        if para_count != 1:
            return None

        if is_list_block_fn(shape):
            return None

        text = tr.Paragraphs(1, 1).Text.strip()
        return text if text else None
    except Exception as e:
        debug_exc_fn("get_single_line_plain_text: 提取单行文本失败", e)
        return None


def split_row_shapes_for_embedding(row_shapes, debug_exc=None, debug_context_prefix="split_row_shapes_for_embedding"):
    """将一行 shape 拆分为普通内容、嵌入 PPT、其他嵌入对象。"""
    normal_shapes = []
    embedded_ppt_shapes = []
    embedded_object_prog_ids = []

    for shape in row_shapes:
        try:
            if shape.Type == 7:  # msoEmbeddedOLEObject
                prog_id = shape.OLEFormat.ProgID
                if "PowerPoint" in str(prog_id):
                    embedded_ppt_shapes.append(shape)
                else:
                    embedded_object_prog_ids.append(str(prog_id))
                continue
        except Exception as e:
            if debug_exc is not None:
                debug_exc(f"{debug_context_prefix}: 检测嵌入对象失败", e)
        normal_shapes.append(shape)

    return normal_shapes, embedded_ppt_shapes, embedded_object_prog_ids


def render_row_text_lines(row_shapes, row_idx, slide_loc, render_shape_row_fn, skip_map=None, image_ctx=None):
    """渲染单行普通 shape 文本。"""
    if not row_shapes:
        return []
    skip_map = skip_map or {}
    slide_loc = str(slide_loc or "").strip()
    row_loc = f"{slide_loc}/R{row_idx}" if slide_loc else None
    return list(
        render_shape_row_fn(
            row_shapes,
            skip_first_para_by_shape_id=skip_map,
            image_ctx=image_ctx,
            loc_prefix=row_loc,
        )
    )


def process_shape_rows_core(
    shape_rows,
    slide_loc,
    row_renderer_fn,
    skip_map=None,
    image_ctx=None,
    embedded_object_line_fn=None,
    debug_exc_fn=None,
    debug_context_prefix="process_shape_rows_core",
):
    """处理 shape 行，返回 (渲染文本行列表, 嵌入PPT shape 列表)。"""
    if row_renderer_fn is None:
        raise ValueError("row_renderer_fn 不能为空")

    debug_exc_fn = debug_exc_fn or (lambda *_args, **_kwargs: None)
    skip_map = skip_map or {}
    rendered_lines = []
    embedded_shapes = []

    for row_idx, row_shapes_in_line in enumerate(shape_rows or [], start=1):
        normal_shapes, row_embedded_shapes, embedded_obj_prog_ids = split_row_shapes_for_embedding(
            row_shapes_in_line,
            debug_exc=debug_exc_fn,
            debug_context_prefix=debug_context_prefix,
        )
        embedded_shapes.extend(row_embedded_shapes)

        for prog_id in embedded_obj_prog_ids:
            if embedded_object_line_fn is None:
                continue
            line = embedded_object_line_fn(prog_id)
            if line is not None:
                rendered_lines.append(str(line))

        row_lines = render_row_text_lines(
            normal_shapes,
            row_idx=row_idx,
            slide_loc=slide_loc,
            render_shape_row_fn=row_renderer_fn,
            skip_map=skip_map,
            image_ctx=image_ctx,
        )
        rendered_lines.extend(row_lines)

    return rendered_lines, embedded_shapes


def render_shape_row_with_number_merge(
    row_shapes,
    skip_first_para_by_shape_id=None,
    image_ctx=None,
    loc_prefix=None,
    safe_shape_id_fn=None,
    get_single_line_plain_text_fn=None,
    escape_md_text_line_fn=None,
    extract_text_from_shape_fn=None,
):
    """渲染一行 shape，支持“编号 shape + 标题 shape”合并。"""
    if not row_shapes:
        return []

    if safe_shape_id_fn is None:
        raise ValueError("safe_shape_id_fn 不能为空")
    if get_single_line_plain_text_fn is None:
        raise ValueError("get_single_line_plain_text_fn 不能为空")
    if escape_md_text_line_fn is None:
        raise ValueError("escape_md_text_line_fn 不能为空")
    if extract_text_from_shape_fn is None:
        raise ValueError("extract_text_from_shape_fn 不能为空")

    merged_lines = []
    used = set()
    skip_first_para_by_shape_id = skip_first_para_by_shape_id or {}

    num_i = None
    num_val = None
    for i, shape in enumerate(row_shapes):
        sid = safe_shape_id_fn(shape)
        if sid is not None and sid in skip_first_para_by_shape_id:
            continue
        text = get_single_line_plain_text_fn(shape)
        if text is None:
            continue
        m = re.fullmatch(r"(\d+)\.?", text)
        if m:
            num_i = i
            num_val = m.group(1)
            break

    title_i = None
    title_text = None
    if num_i is not None:
        for j in range(num_i + 1, len(row_shapes)):
            sid = safe_shape_id_fn(row_shapes[j])
            if sid is not None and sid in skip_first_para_by_shape_id:
                continue
            text = get_single_line_plain_text_fn(row_shapes[j])
            if text is None:
                continue
            if re.fullmatch(r"(\d+)\.?", text):
                continue
            title_i = j
            title_text = text
            break

    if num_i is not None and title_i is not None:
        merged_lines.append(f"{int(num_val)}. {escape_md_text_line_fn(title_text)}")
        used.add(num_i)
        used.add(title_i)

    for i, shape in enumerate(row_shapes):
        if i in used:
            continue
        sid = safe_shape_id_fn(shape)
        skip_text = skip_first_para_by_shape_id.get(sid) if sid is not None else None
        shape_loc = None
        if loc_prefix:
            if sid is not None:
                shape_loc = f"{loc_prefix}/SH{sid}"
            else:
                shape_loc = f"{loc_prefix}/I{i + 1}"

        for text in extract_text_from_shape_fn(
            shape,
            skip_first_para_text=skip_text,
            image_ctx=image_ctx,
            image_loc=shape_loc,
        ):
            merged_lines.append(text)

    return merged_lines


def extract_text_from_shape_core(
    shape,
    skip_first_para_text=None,
    image_ctx=None,
    image_loc=None,
    table_header_mode="first-row",
    export_shape_image_markdown_fn=None,
    debug_exc_fn=None,
    is_list_block_fn=None,
    split_manual_ordered_prefix_fn=None,
    looks_like_brief_list_item_fn=None,
    escape_md_text_line_fn=None,
    strip_bullet_like_prefix_fn=None,
    escape_md_table_cell_fn=None,
):
    """从单个 Shape 提取文本（核心实现）。"""
    texts = []
    debug_exc_fn = debug_exc_fn or (lambda *_args, **_kwargs: None)

    # 检测图片 (msoPicture=13, msoLinkedPicture=11)
    try:
        if shape.Type in (13, 11):
            if export_shape_image_markdown_fn is None:
                raise ValueError("export_shape_image_markdown_fn 不能为空")
            texts.append(export_shape_image_markdown_fn(shape, image_ctx=image_ctx, image_loc=image_loc))
            return texts
    except Exception as e:
        debug_exc_fn("extract_text_from_shape: 读取shape.Type失败", e)

    if shape.HasTextFrame:
        try:
            tr = shape.TextFrame.TextRange
            para_count = tr.Paragraphs().Count

            if para_count == 0:
                return texts

            if is_list_block_fn is None:
                raise ValueError("is_list_block_fn 不能为空")
            is_list = is_list_block_fn(shape)

            # 有些 PPT 的列表是“单层缩进”但启用了项目符号/编号（IndentLevel 都是 1），
            # 此时 is_list_block() 会返回 False；这里探测 Bullet.Visible 来补齐判定。
            has_bullet = False
            try:
                for pi in range(1, para_count + 1):
                    p = tr.Paragraphs(pi, 1)
                    try:
                        if bool(p.ParagraphFormat.Bullet.Visible):
                            has_bullet = True
                            break
                    except Exception:
                        continue
            except Exception:
                has_bullet = False

            # 对编号列表做计数（按 IndentLevel 分组），用于输出 Markdown 有序列表
            ol_counters = {}
            manual_ordered_count = 0
            manual_ordered_base_level = None
            try:
                for pi in range(1, para_count + 1):
                    p = tr.Paragraphs(pi, 1)
                    p_text = p.Text.strip()
                    if split_manual_ordered_prefix_fn is None:
                        raise ValueError("split_manual_ordered_prefix_fn 不能为空")
                    if split_manual_ordered_prefix_fn(p_text) is not None:
                        manual_ordered_count += 1
                        if manual_ordered_base_level is None:
                            try:
                                level = int(p.IndentLevel) - 1
                                if level < 0:
                                    level = 0
                            except Exception:
                                level = 0
                            manual_ordered_base_level = level
            except Exception:
                manual_ordered_count = 0
                manual_ordered_base_level = None
            has_manual_ordered_block = (manual_ordered_count >= 2)

            skip_first_para_text = (str(skip_first_para_text).strip() if skip_first_para_text else None)
            for i in range(1, para_count + 1):
                para = tr.Paragraphs(i, 1)
                text = para.Text.strip()
                if not text:
                    continue
                if skip_first_para_text and i == 1 and text == skip_first_para_text:
                    continue

                # 读取项目符号/编号信息（用于决定无序/有序列表）
                bullet_visible = False
                bullet_type = None
                try:
                    bullet_visible = bool(para.ParagraphFormat.Bullet.Visible)
                except Exception:
                    bullet_visible = False
                try:
                    bullet_type = int(para.ParagraphFormat.Bullet.Type)
                except Exception:
                    bullet_type = None

                if is_list or has_bullet:
                    # 列表格式：根据缩进级别添加前缀
                    # IndentLevel 从 1 开始，转换为 0-based 缩进
                    try:
                        level = int(para.IndentLevel) - 1
                        if level < 0:
                            level = 0
                    except Exception:
                        level = 0
                    indent = "  " * level

                    if not bullet_visible:
                        # 同一 shape 里混排“标题/说明 + 列表”时，非 bullet 段落按普通段落输出
                        # 并重置编号计数，避免跨段污染。
                        ol_counters.clear()
                        if split_manual_ordered_prefix_fn is None:
                            raise ValueError("split_manual_ordered_prefix_fn 不能为空")
                        manual_ol = split_manual_ordered_prefix_fn(text)
                        if manual_ol is not None:
                            if escape_md_text_line_fn is None:
                                raise ValueError("escape_md_text_line_fn 不能为空")
                            n, body = manual_ol
                            texts.append(f"{indent}{n}. {escape_md_text_line_fn(body)}")
                            continue
                        if (has_manual_ordered_block and i > 1 and
                                manual_ordered_base_level is not None and level > int(manual_ordered_base_level)):
                            if looks_like_brief_list_item_fn is None:
                                raise ValueError("looks_like_brief_list_item_fn 不能为空")
                            if looks_like_brief_list_item_fn(text):
                                if escape_md_text_line_fn is None:
                                    raise ValueError("escape_md_text_line_fn 不能为空")
                                texts.append(f"{indent}- {escape_md_text_line_fn(text)}")
                                continue
                        if escape_md_text_line_fn is None:
                            raise ValueError("escape_md_text_line_fn 不能为空")
                        texts.append(escape_md_text_line_fn(text))
                        continue

                    # ppBulletNumbered=2：编号列表（数字在 PPT 格式里，TextRange.Text 不包含“1.”）
                    if bullet_visible and bullet_type == 2:
                        # 清理更深层级计数，避免跨层污染
                        for k in list(ol_counters.keys()):
                            if int(k) > int(level):
                                ol_counters.pop(k, None)

                        start_val = 1
                        try:
                            start_val = int(para.ParagraphFormat.Bullet.StartValue)
                        except Exception:
                            start_val = 1

                        if level not in ol_counters:
                            ol_counters[level] = start_val
                        else:
                            ol_counters[level] = int(ol_counters[level]) + 1
                        n = int(ol_counters[level])
                        if escape_md_text_line_fn is None:
                            raise ValueError("escape_md_text_line_fn 不能为空")
                        texts.append(f"{indent}{n}. {escape_md_text_line_fn(text)}")
                    else:
                        # 无序列表：保持旧行为
                        ol_counters.clear()
                        marker = "*" if is_list else "-"
                        if escape_md_text_line_fn is None:
                            raise ValueError("escape_md_text_line_fn 不能为空")
                        texts.append(f"{indent}{marker} {escape_md_text_line_fn(text)}")
                else:
                    # 普通段落：
                    # 1) 行首“手打符号”（► • ➤ 等）统一转为 "- "
                    # 2) 若该段落在PPT中启用了项目符号(Bullet.Visible)，但脚本未判定为列表块，则也输出为 "- "
                    try:
                        level = int(para.IndentLevel) - 1
                        if level < 0:
                            level = 0
                    except Exception:
                        level = 0
                    indent = "  " * level

                    if split_manual_ordered_prefix_fn is None:
                        raise ValueError("split_manual_ordered_prefix_fn 不能为空")
                    manual_ol = split_manual_ordered_prefix_fn(text)
                    if manual_ol is not None:
                        if escape_md_text_line_fn is None:
                            raise ValueError("escape_md_text_line_fn 不能为空")
                        n, body = manual_ol
                        texts.append(f"{indent}{n}. {escape_md_text_line_fn(body)}")
                        continue

                    if strip_bullet_like_prefix_fn is None:
                        raise ValueError("strip_bullet_like_prefix_fn 不能为空")
                    normalized = strip_bullet_like_prefix_fn(text)
                    if normalized is not None:
                        if escape_md_text_line_fn is None:
                            raise ValueError("escape_md_text_line_fn 不能为空")
                        texts.append(f"{indent}- {escape_md_text_line_fn(normalized)}")
                        continue

                    if (has_manual_ordered_block and i > 1 and
                            manual_ordered_base_level is not None and level > int(manual_ordered_base_level)):
                        if looks_like_brief_list_item_fn is None:
                            raise ValueError("looks_like_brief_list_item_fn 不能为空")
                        if looks_like_brief_list_item_fn(text):
                            if escape_md_text_line_fn is None:
                                raise ValueError("escape_md_text_line_fn 不能为空")
                            texts.append(f"{indent}- {escape_md_text_line_fn(text)}")
                            continue

                    if escape_md_text_line_fn is None:
                        raise ValueError("escape_md_text_line_fn 不能为空")
                    texts.append(escape_md_text_line_fn(text))
        except Exception as e:
            debug_exc_fn("extract_text_from_shape: 解析TextFrame失败，尝试回退", e)
            # 回退到原始方式
            try:
                text = shape.TextFrame.TextRange.Text
                if text and text.strip():
                    if escape_md_text_line_fn is None:
                        raise ValueError("escape_md_text_line_fn 不能为空")
                    texts.append(escape_md_text_line_fn(text))
            except Exception as e:
                debug_exc_fn("extract_text_from_shape: 回退读取TextRange.Text失败", e)

    # 处理表格
    try:
        if shape.HasTable:
            table = shape.Table
            rows = []
            for r in range(1, table.Rows.Count + 1):
                row = []
                for c in range(1, table.Columns.Count + 1):
                    try:
                        cell = table.Cell(r, c).Shape.TextFrame.TextRange.Text.strip()
                        if escape_md_table_cell_fn is None:
                            raise ValueError("escape_md_table_cell_fn 不能为空")
                        row.append(escape_md_table_cell_fn(cell))
                    except Exception as e:
                        debug_exc_fn("extract_text_from_shape: 读取表格单元格失败", e)
                        row.append("")
                rows.append(row)

            if rows:
                # Markdown表格
                col_count = len(rows[0]) if rows[0] else 0
                if col_count == 0:
                    return texts

                if str(table_header_mode) == "empty":
                    header = [""] * col_count
                    md = "| " + " | ".join(header) + " |\n"
                    md += "| " + " | ".join(["---"] * col_count) + " |\n"
                    for row in rows:
                        md += "| " + " | ".join(row) + " |\n"
                else:
                    md = "| " + " | ".join(rows[0]) + " |\n"
                    md += "| " + " | ".join(["---"] * col_count) + " |\n"
                    for row in rows[1:]:
                        md += "| " + " | ".join(row) + " |\n"
                texts.append(md)
    except Exception as e:
        debug_exc_fn("extract_text_from_shape: 处理表格失败", e)

    return texts
