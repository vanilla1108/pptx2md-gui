"""Markdown 渲染辅助函数。"""


def format_loc(loc_parts):
    """把位置路径格式化为可读且可搜索的串，如：S2/E1/S1。"""
    if not loc_parts:
        return ""
    return "/".join([str(x) for x in loc_parts if x])


def sanitize_md_comment_text(text):
    """清洗注释文本，避免破坏 HTML 注释结构。"""
    s = str(text or "")
    s = s.replace("\r", " ").replace("\n", " ").strip()
    s = s.replace("-->", "--＞")
    return s


def md_comment(text):
    """生成 HTML 注释行（含换行）。"""
    s = sanitize_md_comment_text(text)
    if not s:
        return ""
    return f"<!-- {s} -->\n"


def md_heading(level, title):
    """生成 Markdown 标题行（含末尾空行）。"""
    level = int(level)
    if level < 1:
        level = 1
    if level > 6:
        level = 6
    return f"{'#' * level} {title}\n\n"


def md_path_quote(loc):
    """生成路径注释行（不含额外空行）。"""
    loc = str(loc or "").strip()
    if not loc:
        return ""
    return md_comment(f"path: {loc}")


def md_heading_with_path(level, title, loc):
    """生成 Markdown 标题 + 下一行路径注释（无空行间隔）。"""
    level = int(level)
    if level < 1:
        level = 1
    if level > 6:
        level = 6
    heading = f"{'#' * level} {title}\n"
    path_line = md_path_quote(loc)
    if path_line:
        return heading + path_line + "\n"
    return heading + "\n"


def md_embedded_ppt_marker(title, loc):
    """生成嵌入 PPT 对象的单行注释标记。"""
    title = str(title or "").strip()
    loc = str(loc or "").strip()
    label = f"embedded-ppt: {title}" if title else "embedded-ppt"
    if loc:
        return md_comment(f"{label} | path: {loc}")
    return md_comment(label)


def md_hr():
    """Markdown 分隔线。"""
    return "---\n\n"


def md_slide_heading_with_ref(level, title, slide_label, slide_no, loc):
    """生成“注释标记 + 幻灯片标题”块。"""
    level = int(level)
    if level < 1:
        level = 1
    if level > 6:
        level = 6

    title = str(title or "").replace("\r", " ").replace("\n", " ").strip()
    if title == "":
        title = f"{slide_label} {slide_no}"

    loc = str(loc or "").strip()
    kind = "slide" if str(slide_label) == "幻灯片" else "embedded-slide"
    comment = f"{kind}: {slide_no}"
    if loc:
        comment += f" | path: {loc}"

    return f"{md_comment(comment)}{'#' * level} {title}\n\n"
