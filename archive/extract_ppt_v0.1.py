# -*- coding: utf-8 -*-
"""
PPT内容提取工具 - 支持嵌套PPT + 视觉位置排序
使用win32com调用本地PowerPoint提取所有文本内容，保存为Markdown格式
文本块按视觉位置排序（从上到下、从左到右）
"""

import win32com.client
import os
import sys
import time
import argparse
import traceback
import re

# 修复Windows控制台编码问题
if sys.platform == 'win32':
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    sys.stderr.reconfigure(encoding='utf-8', errors='replace')

# 同一行判定阈值（磅）
# - "auto": 自适应模式，基于文本高度动态计算（推荐）
# - 数字: 固定阈值，如 22 磅 ≈ 0.3 英寸
ROW_THRESHOLD_POINTS = "auto"
ROW_THRESHOLD_FALLBACK = 22  # 自适应模式的回退值

DEBUG = False
COM_POLL_INTERVAL_SEC = 0.1
COM_OPEN_TIMEOUT_SEC = 15
COM_EMBED_TIMEOUT_SEC = 10
TABLE_HEADER_MODE = "first-row"  # first-row | empty

# 标题层级策略：统一使用二级标题（嵌入层级用"路径"体现即可）
TITLE_HEADING_LEVEL = 2

# XY-Cut 多栏检测配置
VERTICAL_GAP_RATIO = 0.08       # 垂直间隙阈值 = 区域宽度 × 此比例
HORIZONTAL_GAP_RATIO = 0.06     # 水平间隙阈值 = 区域高度 × 此比例
MIN_V_GAP_POINTS = 40.0         # 垂直间隙阈值下限（磅）
MIN_H_GAP_POINTS = 24.0         # 水平间隙阈值下限（磅）
WIDE_SPAN_RATIO = 0.8           # 宽度跨区桥接判定比例
TALL_SPAN_RATIO = 0.9           # 高度跨区桥接判定比例
XY_CUT_MAX_DEPTH = 2            # 最大递归深度
MIN_SHAPES_PER_REGION = 2       # 每个区域最少 shapes 数
XY_CUT_EPS = 0.5                # 浮点容差（磅）


def _format_loc(loc_parts):
    """把位置路径格式化为可读且可搜索的串，如：S2/E1/S1。"""
    if not loc_parts:
        return ""
    return "/".join([str(x) for x in loc_parts if x])


def _md_heading(level, title):
    """生成Markdown标题行（含末尾空行）。"""
    level = int(level)
    if level < 1:
        level = 1
    if level > 6:
        level = 6
    return f"{'#' * level} {title}\n\n"


def _md_path_quote(loc):
    """生成路径引用块行（不含额外空行）。"""
    loc = str(loc or "").strip()
    if not loc:
        return ""
    return f"> `路径：{loc}`\n"


def _md_heading_with_path(level, title, loc):
    """生成Markdown标题 + 下一行路径引用块（无空行间隔）。"""
    level = int(level)
    if level < 1:
        level = 1
    if level > 6:
        level = 6

    heading = f"{'#' * level} {title}\n"
    path_line = _md_path_quote(loc)
    if path_line:
        return heading + path_line + "\n"
    return heading + "\n"


def _md_embedded_ppt_marker(title, loc):
    """生成嵌入PPT对象的单行引用块标记。

    期望输出类似：
    > ▶ 嵌入PPT #1 `路径：S2/E1`
    """
    title = str(title or "").strip()
    loc = str(loc or "").strip()
    if loc:
        return f"> ▶ {title} `路径：{loc}`\n"
    return f"> ▶ {title}\n"


def _md_hr():
    """Markdown分隔线（水平线）。"""
    return "---\n\n"


def _md_slide_heading_with_ref(level, title, slide_label, slide_no, loc):
    """生成“幻灯片标题 + 引用块(含页码与路径)”。

    期望输出类似：
    ## 第五-0讲 ARM体系结构
    > 幻灯片 2：`路径：S2`
    """
    level = int(level)
    if level < 1:
        level = 1
    if level > 6:
        level = 6

    title = str(title or "").replace("\r", " ").replace("\n", " ").strip()
    if title == "":
        title = f"{slide_label} {slide_no}"

    loc = str(loc or "").strip()
    if loc:
        ref = f"> {slide_label} {slide_no}：`路径：{loc}`\n"
    else:
        ref = f"> {slide_label} {slide_no}\n"

    return f"{'#' * level} {title}\n{ref}\n"


def _safe_shape_id(shape):
    try:
        return int(shape.Id)
    except Exception:
        return None


def _first_paragraph_text(shape):
    """读取shape第一段文本（纯文本，strip后返回）。"""
    try:
        if not shape.HasTextFrame or not shape.TextFrame.HasText:
            return None
        tr = shape.TextFrame.TextRange
        if tr.Paragraphs().Count <= 0:
            return None
        return tr.Paragraphs(1, 1).Text.strip() or None
    except Exception as e:
        _debug_exc("_first_paragraph_text: 读取失败", e)
        return None


def _is_title_candidate_shape(shape):
    """判断shape是否可能是“标题行”候选（不保证一定是标题）。"""
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

        if is_list_block(shape):
            return False

        tr = shape.TextFrame.TextRange
        if tr.Paragraphs().Count <= 0:
            return False

        para1 = tr.Paragraphs(1, 1)
        text = para1.Text.strip()
        if not text:
            return False

        # 项目符号行通常不是页标题
        try:
            if bool(para1.ParagraphFormat.Bullet.Visible):
                return False
        except Exception:
            pass

        # 纯数字一般是页码/编号，不当标题
        if re.fullmatch(r"\d+", text):
            return False

        # 过长更像正文段落
        if len(text) > 120:
            return False

        return True
    except Exception as e:
        _debug_exc("_is_title_candidate_shape: 检测失败", e)
        return False


def detect_slide_title(slide):
    """识别单页幻灯片的第一行标题。

    优先使用 slide.Shapes.Title；否则从文本shape中按“靠上 + 字号大 + 文本短”择优。
    返回：
        {
          "title": str|None,
          "shape_id": int|None,   # 标题所在shape id（用于正文中跳过重复输出）
        }
    """
    # 1) 优先 Title 占位符
    try:
        title_shape = slide.Shapes.Title
        title_text = _first_paragraph_text(title_shape)
        if title_text:
            return {"title": title_text, "shape_id": _safe_shape_id(title_shape)}
    except Exception:
        pass

    # 2) 回退：扫描候选shape并打分
    best = None
    try:
        shapes = list(slide.Shapes)
    except Exception:
        shapes = []

    for shape in shapes:
        if not _is_title_candidate_shape(shape):
            continue

        text = _first_paragraph_text(shape)
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

        # 越靠上、字号越大、越短 => 分越高
        score = size * 10.0 - top / 5.0 - len(text) * 0.5
        if top <= 120:
            score += 15.0

        if best is None or score > best["score"]:
            best = {"score": score, "text": text, "shape_id": _safe_shape_id(shape)}

    if best is not None:
        return {"title": best["text"], "shape_id": best["shape_id"]}

    return {"title": None, "shape_id": None}


def _find_shape_by_id_in_slide(slide, shape_id):
    """在 slide.Shapes 中按 shape.Id 查找 Shape（失败则返回 None）。"""
    if shape_id is None:
        return None
    try:
        target = int(shape_id)
    except Exception:
        return None

    try:
        for s in list(slide.Shapes):
            if _safe_shape_id(s) == target:
                return s
    except Exception as e:
        _debug_exc("_find_shape_by_id_in_slide: 枚举Shapes失败", e)
    return None


def _extract_title_shape_extra_lines(slide, title_info):
    """补回“标题shape里除第一段标题外的其他段落内容”。

    说明：
    - 当前实现会在布局分析中剔除 title_shape_id（避免影响 XY-Cut 多栏切割）
    - 但部分PPT把“标题 + 副标题/要点/正文”放在同一个文本框（同一个 shape）
      仅剔除会导致副标题/要点丢失（例如“本节提要”、课程目标列表等）
    - 因此这里把 title_shape 的“除第一段标题外”的内容补回到正文输出中
    """
    try:
        if not title_info:
            return []
        sid = title_info.get("shape_id")
        title = title_info.get("title")
        if not sid or not title:
            return []

        shape = _find_shape_by_id_in_slide(slide, sid)
        if shape is None:
            return []

        lines = extract_text_from_shape(shape, skip_first_para_text=str(title))

        # 防御：避免把标题行再次输出（例如首段匹配失败的极端情况）
        title_norm = str(title).strip()
        out = []
        for t in lines:
            if str(t).strip() == title_norm:
                continue
            out.append(t)
        return out
    except Exception as e:
        _debug_exc("_extract_title_shape_extra_lines: 提取失败", e)
        return []


def _format_exc(e, limit=120):
    msg = str(e) if e is not None else ""
    msg = msg.replace("\r", " ").replace("\n", " ")
    if len(msg) > limit:
        msg = msg[:limit] + "..."
    return f"{type(e).__name__}: {msg}" if msg else type(e).__name__


def _debug_exc(context, e):
    if not DEBUG:
        return
    print(f"[DEBUG] {context}: {_format_exc(e)}", file=sys.stderr)
    # 在非 except 作用域内调用 traceback.print_exc() 会丢失原始异常堆栈；
    # 这里优先打印传入异常对象自带的 traceback，确保定位有效。
    try:
        tb = getattr(e, "__traceback__", None)
        if tb is not None:
            traceback.print_exception(type(e), e, tb)
        else:
            traceback.print_exc()
    except Exception:
        traceback.print_exc()


def _wait_com(action, timeout_sec, context):
    """轮询等待COM调用成功；超时则抛出最后一次异常。"""
    deadline = time.time() + timeout_sec
    start = time.time()
    attempts = 0
    last_exc = None
    first_exc = None
    while time.time() <= deadline:
        attempts += 1
        try:
            result = action()
            # 仅在DEBUG时输出“经历重试后成功”的信息，避免污染正常输出。
            if DEBUG and attempts > 1:
                elapsed = time.time() - start
                print(f"[DEBUG] {context}: 成功(重试{attempts - 1}次, {elapsed:.2f}s)", file=sys.stderr)
            return result
        except Exception as e:
            if first_exc is None:
                first_exc = e
                if DEBUG:
                    print(f"[DEBUG] {context}: 首次失败: {_format_exc(e)}", file=sys.stderr)
            last_exc = e
            time.sleep(COM_POLL_INTERVAL_SEC)
    if last_exc is not None:
        _debug_exc(f"{context}: 超时({timeout_sec}s, 尝试{attempts}次)", last_exc)
        raise last_exc
    raise TimeoutError(f"{context}: 超时({timeout_sec}s)")


def _try_call(action, context):
    try:
        return action()
    except Exception as e:
        _debug_exc(context, e)
        return None


def _close_embedded_object(ppt_app, embedded_ppt=None):
    """尽力退出嵌入对象编辑态，避免状态污染影响后续处理。"""
    if embedded_ppt is not None:
        _try_call(lambda: embedded_ppt.Close(), "close_embedded_object: embedded_ppt.Close失败")
    _try_call(lambda: ppt_app.CommandBars.ExecuteMso('ObjectCloseAndReturn'),
              "close_embedded_object: ExecuteMso(ObjectCloseAndReturn)失败")
    _try_call(lambda: setattr(ppt_app.ActiveWindow, "ViewType", 1),
              "close_embedded_object: 强制切回Normal View失败")


def _escape_md_text_line(text):
    """转义会破坏Markdown结构的常见情况（用于普通段落行）。"""
    if text is None:
        return ""
    # 先做基本规范化：避免把CR/LF带入Markdown结构
    text = str(text)
    text = text.replace("\r", " ")
    text = text.replace("\n", " ")
    text = text.replace("\x0b", " ")
    text = text.replace("\t", " ")

    # 保留原始缩进（有些PPT文本会带前导空格）
    stripped = text.lstrip(" ")
    prefix = text[:len(text) - len(stripped)]

    if stripped == "":
        return text

    # 防止误识别为标题/引用/列表/分割线
    if stripped.startswith(("#", ">")):
        stripped = "\\" + stripped
    elif stripped.startswith(("-", "*", "+")) and len(stripped) >= 2 and stripped[1] == " ":
        stripped = "\\" + stripped
    elif re.match(r"^\d+\.\s", stripped):
        stripped = "\\" + stripped
    elif re.match(r"^(-{3,}|\*{3,}|_{3,})$", stripped.strip()):
        stripped = "\\" + stripped

    return prefix + stripped


def _escape_md_table_cell(text):
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


_BULLET_LIKE_PREFIXES = [
    # 常见“手打/装饰”行首符号（字符本身出现在文本里）
    "►", "▸", "▹", "▷", "▶", "▻",
    "•", "◦", "‣", "∙", "·", "●", "○",
    "◆", "◇", "■", "□",
    "➤", "➢", "➣", "➔", "→",
    "➜", "➙", "➛", "➝", "➞", "➟",
]

_BULLET_LIKE_PREFIX_RE = re.compile(
    r"^(?:%s)[\s\u00a0]*" % "|".join([re.escape(x) for x in _BULLET_LIKE_PREFIXES])
)


def _strip_bullet_like_prefix(text):
    """若文本以常见“项目符号样式字符”开头，则去掉该符号并返回剩余正文；否则返回None。

    仅用于普通段落（非脚本自身输出的引用/提示行）。
    """
    if text is None:
        return None
    s = str(text)
    # 规范化：避免控制符影响匹配
    s = s.replace("\r", " ").replace("\n", " ").replace("\x0b", " ").replace("\t", " ")
    s = s.strip()
    if not s:
        return None

    m = _BULLET_LIKE_PREFIX_RE.match(s)
    if not m:
        return None

    rest = s[m.end():].lstrip(" ")
    return rest if rest else None


def get_unique_output_path(base_path):
    """获取唯一的输出路径，如果文件已存在则添加序号

    Args:
        base_path: 基础文件路径

    Returns:
        不冲突的文件路径
    """
    if not os.path.exists(base_path):
        return base_path

    dir_name = os.path.dirname(base_path)
    base_name = os.path.splitext(os.path.basename(base_path))[0]
    ext = os.path.splitext(base_path)[1]

    counter = 1
    while True:
        new_path = os.path.join(dir_name, f"{base_name}_{counter}{ext}")
        if not os.path.exists(new_path):
            return new_path
        counter += 1


def sort_shapes_by_visual_position(shapes, row_threshold_points="auto",
                                   enable_xy_cut=False, slide_size=None,
                                   exclude_shape_ids=None):
    """按视觉位置排序形状（支持行分组，可选 XY-Cut 多栏感知）

    当形状的 top 值差距小于阈值时，认为它们在同一行，
    同一行内按 left（从左到右）排序。

    Args:
        shapes: 形状集合
        row_threshold_points: 同行判定阈值
            - "auto": 自适应模式，基于文本高度动态计算（推荐）
            - 数字: 固定阈值（磅）
        enable_xy_cut: 是否启用 XY-Cut 多栏检测
        slide_size: (width, height) 幻灯片尺寸（磅）
        exclude_shape_ids: set[int]|None，排除的 shape id 集合

    Returns:
        排序后的形状列表
    """
    rows = group_shapes_by_visual_rows(
        shapes,
        row_threshold_points=row_threshold_points,
        enable_xy_cut=enable_xy_cut,
        slide_size=slide_size,
        exclude_shape_ids=exclude_shape_ids,
    )
    return [s for row in rows for s in row]


def _compute_adaptive_row_threshold(shapes):
    """根据 shapes 的文本高度动态计算行判定阈值。

    策略：收集所有文本框的实际文本高度，取中位数的 1.3 倍作为阈值。
    这样可以适应不同字号的 PPT，避免固定阈值导致的误分行/误合行。

    Args:
        shapes: 形状集合

    Returns:
        float: 计算出的阈值（磅），若无法计算则返回 ROW_THRESHOLD_FALLBACK
    """
    heights = []
    for shape in shapes:
        try:
            if not shape.HasTextFrame:
                continue
            if not shape.TextFrame.HasText:
                continue
            tr = shape.TextFrame.TextRange
            bound_height = float(tr.BoundHeight)
            if bound_height > 0:
                heights.append(bound_height)
        except Exception:
            pass

    if not heights:
        return ROW_THRESHOLD_FALLBACK

    # 使用中位数，对异常值更鲁棒
    heights.sort()
    median_height = heights[len(heights) // 2]

    # 阈值 = 中位数高度 × 1.3（经验系数，允许轻微错位）
    threshold = median_height * 1.3

    # 设置合理的上下限，避免极端情况
    threshold = max(threshold, 10.0)   # 最小 10 磅
    threshold = min(threshold, 100.0)  # 最大 100 磅

    if DEBUG:
        print(f"[DEBUG] 自适应阈值: 中位数高度={median_height:.1f}磅, 阈值={threshold:.1f}磅 (样本数={len(heights)})")

    return threshold


def _shape_anchor_xy(shape):
    """返回用于视觉排序的锚点坐标(x, y)（单位：磅）。

    经验上，Shape.Top/Left 是外框；对于不同字号/不同内边距的文本框，
    仅用Top会把"看起来同一行"的内容拆成两行。

    这里尽量使用TextRange的Bound*（文本实际包围盒）来估计锚点。
    """
    try:
        # 默认回退到形状外框中心
        x = float(shape.Left) + float(shape.Width) / 2.0
        y = float(shape.Top) + float(shape.Height) / 2.0
    except Exception as e:
        _debug_exc("_shape_anchor_xy: 读取Shape.Left/Top/Width/Height失败", e)
        return float("inf"), float("inf")

    try:
        if shape.HasTextFrame and shape.TextFrame.HasText:
            tr = shape.TextFrame.TextRange
            # Bound* 是文本实际包围盒相对TextFrame的位置（实践中一般可加上shape.Top/Left）
            bx = float(shape.Left) + float(tr.BoundLeft) + float(tr.BoundWidth) / 2.0
            by = float(shape.Top) + float(tr.BoundTop) + float(tr.BoundHeight) / 2.0
            # 只要能读到Bound*，就用它（比外框更接近视觉位置）
            x, y = bx, by
    except Exception as e:
        _debug_exc("_shape_anchor_xy: 读取TextRange.Bound*失败，回退到外框中心", e)

    return x, y


def _shape_bbox(shape):
    """返回 shape 的边界框 (left, top, right, bottom)，单位：磅

    策略：
    - 优先使用 TextRange.Bound*（文本实际包围盒）
    - Clamp 到 shape 外框范围内
    - 回退到 shape 外框
    - 返回 None 表示读取失败
    """
    # 1. 读取 shape 外框
    try:
        left = float(shape.Left)
        top = float(shape.Top)
        right = left + float(shape.Width)
        bottom = top + float(shape.Height)
    except Exception:
        return None

    # 2. 尝试使用文本包围盒
    try:
        if shape.HasTextFrame and shape.TextFrame.HasText:
            tr = shape.TextFrame.TextRange
            bw = float(tr.BoundWidth)
            bh = float(tr.BoundHeight)

            if bw > 0 and bh > 0:
                text_left = left + float(tr.BoundLeft)
                text_top = top + float(tr.BoundTop)
                text_right = text_left + bw
                text_bottom = text_top + bh

                # Clamp 到 shape 外框
                text_left = max(text_left, left)
                text_top = max(text_top, top)
                text_right = min(text_right, right)
                text_bottom = min(text_bottom, bottom)

                if text_right > text_left and text_bottom > text_top:
                    return (text_left, text_top, text_right, text_bottom)
    except Exception:
        pass

    return (left, top, right, bottom)


def _is_wide_shape(bbox, region_bbox):
    """宽度跨区桥接：用于垂直切割时忽略标题/页眉/页脚等"满宽"元素。"""
    bw = bbox[2] - bbox[0]
    rw = region_bbox[2] - region_bbox[0]
    return rw > 0 and (bw / rw) >= WIDE_SPAN_RATIO


def _is_tall_shape(bbox, region_bbox):
    """高度跨区桥接：用于水平切割时忽略少量"满高"元素（如侧边装饰条）。"""
    bh = bbox[3] - bbox[1]
    rh = region_bbox[3] - region_bbox[1]
    return rh > 0 and (bh / rh) >= TALL_SPAN_RATIO


def _find_vertical_cut(boxes, region_bbox):
    """寻找有效的垂直切割点

    Args:
        boxes: [(shape, bbox), ...] 候选 shapes 及其 bbox
        region_bbox: 当前区域边界框

    Returns:
        cut_x: 切割点 x 坐标，若无有效切割则返回 None
    """
    # 1. 过滤跨栏桥接 shape
    narrow_boxes = [(s, b) for s, b in boxes if not _is_wide_shape(b, region_bbox)]

    if len(narrow_boxes) < 2:
        return None

    # 2. 计算阈值
    region_width = region_bbox[2] - region_bbox[0]
    gap_threshold = max(region_width * VERTICAL_GAP_RATIO, MIN_V_GAP_POINTS)

    # 3. 合并区间方式寻找最大间隙
    sorted_boxes = sorted(narrow_boxes, key=lambda x: x[1][0])

    max_gap = 0
    best_cut = None
    right_edge = sorted_boxes[0][1][2]

    for i in range(1, len(sorted_boxes)):
        current_left = sorted_boxes[i][1][0]
        gap = current_left - right_edge

        if gap > max_gap and gap >= gap_threshold:
            max_gap = gap
            best_cut = (right_edge + current_left) / 2.0

        right_edge = max(right_edge, sorted_boxes[i][1][2])

    if best_cut is None:
        return None

    # 4. 校验无跨越
    for shape, bbox in narrow_boxes:
        if bbox[0] < best_cut - XY_CUT_EPS and best_cut + XY_CUT_EPS < bbox[2]:
            return None

    # 5. 校验两侧数量
    left_count = sum(1 for _, b in narrow_boxes if b[2] <= best_cut + XY_CUT_EPS)
    right_count = sum(1 for _, b in narrow_boxes if b[0] >= best_cut - XY_CUT_EPS)

    if left_count < MIN_SHAPES_PER_REGION or right_count < MIN_SHAPES_PER_REGION:
        return None

    return best_cut


def _find_horizontal_cut(boxes, region_bbox):
    """寻找有效的水平切割点（用于"标题带 + 两栏正文"的第一刀）

    思路与 _find_vertical_cut 对称：按 top 排序，寻找最大垂直间隙。
    """
    # 1. 过滤跨行桥接 shape（满高装饰条等）
    short_boxes = [(s, b) for s, b in boxes if not _is_tall_shape(b, region_bbox)]
    if len(short_boxes) < 2:
        return None

    # 2. 计算阈值
    region_height = region_bbox[3] - region_bbox[1]
    gap_threshold = max(region_height * HORIZONTAL_GAP_RATIO, MIN_H_GAP_POINTS)

    # 3. 合并区间方式寻找最大间隙（按 top）
    sorted_boxes = sorted(short_boxes, key=lambda x: x[1][1])

    max_gap = 0
    best_cut = None
    bottom_edge = sorted_boxes[0][1][3]

    for i in range(1, len(sorted_boxes)):
        current_top = sorted_boxes[i][1][1]
        gap = current_top - bottom_edge
        if gap > max_gap and gap >= gap_threshold:
            max_gap = gap
            best_cut = (bottom_edge + current_top) / 2.0
        bottom_edge = max(bottom_edge, sorted_boxes[i][1][3])

    if best_cut is None:
        return None

    # 4. 校验无跨越
    for _, bbox in short_boxes:
        if bbox[1] < best_cut - XY_CUT_EPS and best_cut + XY_CUT_EPS < bbox[3]:
            return None

    # 5. 校验两侧数量
    top_count = sum(1 for _, b in short_boxes if b[3] <= best_cut + XY_CUT_EPS)
    bottom_count = sum(1 for _, b in short_boxes if b[1] >= best_cut - XY_CUT_EPS)
    if top_count < MIN_SHAPES_PER_REGION or bottom_count < MIN_SHAPES_PER_REGION:
        return None

    return best_cut


def _xy_cut_partition(shapes, region_bbox, depth, bbox_cache, row_threshold_points):
    """递归 XY-Cut 分区

    Args:
        shapes: 待分区的 shape 列表
        region_bbox: 当前区域边界框 (left, top, right, bottom)
        depth: 当前递归深度
        bbox_cache: {id(shape): bbox} 缓存
        row_threshold_points: 行阈值参数

    Returns:
        List[List[shape]]: 按阅读顺序排列的区域列表（每个区域是 shape 列表）
    """
    # 构造 boxes（跳过 None）
    boxes = []
    for shape in shapes:
        bbox = bbox_cache.get(id(shape))
        if bbox is None:
            # 不允许静默丢 shape：只要该区域出现无法读 bbox 的元素，就回退到旧算法
            fallback_rows = _group_shapes_by_visual_rows_simple(
                list(shapes), row_threshold_points
            )
            return [[s for row in fallback_rows for s in row]]
        boxes.append((shape, bbox))

    if len(boxes) == 0:
        return []
    if len(boxes) == 1:
        return [[boxes[0][0]]]

    # 达到最大深度
    if depth >= XY_CUT_MAX_DEPTH:
        sorted_rows = _group_shapes_by_visual_rows_simple(
            [s for s, _ in boxes], row_threshold_points
        )
        return [[s for row in sorted_rows for s in row]]

    # 决定切割方向
    region_width = region_bbox[2] - region_bbox[0]
    region_height = region_bbox[3] - region_bbox[1]
    # 经验规则：
    # - 顶层若存在"满宽且靠上"的元素，更可能是"标题带 + 正文"，优先水平切
    # - 否则按长宽比决定（宽页优先垂直切，减少双栏被水平误切的概率）
    has_top_wide = any(
        _is_wide_shape(b, region_bbox) and (b[1] - region_bbox[1]) <= region_height * 0.25
        for _, b in boxes
    )
    prefer_vertical = (not has_top_wide) and (region_width > region_height * 1.5)

    h_cut = _find_horizontal_cut(boxes, region_bbox)
    v_cut = _find_vertical_cut(boxes, region_bbox)

    chosen_cut = None
    is_horizontal = False

    if prefer_vertical:
        if v_cut is not None:
            chosen_cut, is_horizontal = v_cut, False
        elif h_cut is not None:
            chosen_cut, is_horizontal = h_cut, True
    else:
        if h_cut is not None:
            chosen_cut, is_horizontal = h_cut, True
        elif v_cut is not None:
            chosen_cut, is_horizontal = v_cut, False

    # 无法切割
    if chosen_cut is None:
        sorted_rows = _group_shapes_by_visual_rows_simple(
            [s for s, _ in boxes], row_threshold_points
        )
        return [[s for row in sorted_rows for s in row]]

    # 按中心点分配
    first_shapes = []
    second_shapes = []

    for s, b in boxes:
        if is_horizontal:
            center = (b[1] + b[3]) / 2.0
        else:
            center = (b[0] + b[2]) / 2.0

        if center < chosen_cut:
            first_shapes.append(s)
        else:
            second_shapes.append(s)

    # 计算子区域 bbox
    if is_horizontal:
        first_bbox = (region_bbox[0], region_bbox[1], region_bbox[2], chosen_cut)
        second_bbox = (region_bbox[0], chosen_cut, region_bbox[2], region_bbox[3])
    else:
        first_bbox = (region_bbox[0], region_bbox[1], chosen_cut, region_bbox[3])
        second_bbox = (chosen_cut, region_bbox[1], region_bbox[2], region_bbox[3])

    # 递归
    result = []
    result.extend(_xy_cut_partition(first_shapes, first_bbox, depth + 1,
                                    bbox_cache, row_threshold_points))
    result.extend(_xy_cut_partition(second_shapes, second_bbox, depth + 1,
                                    bbox_cache, row_threshold_points))

    return result


def _compute_region_bbox_from_cache(shapes, cache):
    """从 bbox_cache 推断当前 shapes 的包围框。

    注意：若存在 None（bbox 读取失败）应返回 None 触发整体回退。
    """
    lefts, tops, rights, bottoms = [], [], [], []
    for s in shapes:
        b = cache.get(id(s))
        if b is None:
            return None
        lefts.append(b[0])
        tops.append(b[1])
        rights.append(b[2])
        bottoms.append(b[3])
    if not lefts:
        return None
    return (min(lefts), min(tops), max(rights), max(bottoms))


def _xy_cut_regions_to_rows(regions, row_threshold_points, slide_size=None):
    """将 XY-Cut 的 regions 输出转换为 rows。

    原则：region 内排序完全复用旧的行分组逻辑，从而保持
    `_render_texts_from_shape_row()` 的"编号shape+标题shape"合并行为。
    """
    rows = []
    for region_shapes in regions:
        region_rows = _group_shapes_by_visual_rows_simple(
            region_shapes, row_threshold_points
        )
        rows.extend(region_rows)
    return rows


def _group_shapes_by_visual_rows_simple(shapes, row_threshold_points="auto"):
    """按视觉位置对shape分行并在行内排序（返回二维数组）- 简单版本，不含 XY-Cut。

    - 行判定：按锚点y从上到下；y差小于阈值视为同一行
    - 行内排序：按锚点x从左到右

    Args:
        shapes: 形状集合
        row_threshold_points: 同行判定阈值
            - "auto": 自适应模式，基于文本高度动态计算（推荐）
            - 数字: 固定阈值（磅）

    Returns:
        List[List[shape]]
    """
    shapes_list = list(shapes)

    # 解析阈值参数
    if row_threshold_points == "auto":
        threshold = _compute_adaptive_row_threshold(shapes_list)
    else:
        threshold = float(row_threshold_points)

    items = []
    for shape in shapes_list:
        try:
            x, y = _shape_anchor_xy(shape)
            items.append({"x": x, "y": y, "shape": shape})
        except Exception as e:
            _debug_exc("_group_shapes_by_visual_rows_simple: 读取锚点失败", e)
            items.append({"x": float("inf"), "y": float("inf"), "shape": shape})

    items.sort(key=lambda it: (it["y"], it["x"]))

    rows = []
    current = []
    current_y = None

    for it in items:
        if current_y is None:
            current = [it]
            current_y = it["y"]
            continue

        if abs(it["y"] - current_y) <= threshold:
            current.append(it)
            # 动态更新本行中心y，避免"链式接近"导致误切行
            current_y = (current_y * (len(current) - 1) + it["y"]) / float(len(current))
        else:
            current.sort(key=lambda x: x["x"])
            rows.append([x["shape"] for x in current])
            current = [it]
            current_y = it["y"]

    if current:
        current.sort(key=lambda x: x["x"])
        rows.append([x["shape"] for x in current])

    return rows


def group_shapes_by_visual_rows(shapes, row_threshold_points="auto",
                                enable_xy_cut=False, slide_size=None,
                                exclude_shape_ids=None):
    """按视觉位置对 shape 分行并在行内排序（可选：XY-Cut 多栏感知）。

    - enable_xy_cut=False：保持现有行为（仅锚点排序+行分组）
    - enable_xy_cut=True：先做 XY-Cut 得到 regions，再对每个 region 回退到旧算法产出 rows

    Args:
        shapes: 形状集合
        row_threshold_points: 同行判定阈值
            - "auto": 自适应模式，基于文本高度动态计算（推荐）
            - 数字: 固定阈值（磅）
        enable_xy_cut: 是否启用 XY-Cut 多栏检测
        slide_size: (width, height) 幻灯片尺寸（磅），用于计算 region_bbox
        exclude_shape_ids: set[int]|None，调用方可用它排除已单独渲染的标题 shape

    Returns:
        List[List[shape]]
    """
    shapes_list = list(shapes)

    # 排除指定的 shape（如已单独渲染的标题）
    if exclude_shape_ids:
        shapes_list = [s for s in shapes_list if _safe_shape_id(s) not in exclude_shape_ids]

    # 不启用 XY-Cut 或 shape 数量太少，直接使用简单版本
    if not enable_xy_cut or len(shapes_list) <= 2:
        return _group_shapes_by_visual_rows_simple(shapes_list, row_threshold_points)

    # 预计算 bbox_cache
    bbox_cache = {id(s): _shape_bbox(s) for s in shapes_list}

    # 计算 region_bbox
    if slide_size is not None:
        region_bbox = (0, 0, slide_size[0], slide_size[1])
    else:
        region_bbox = _compute_region_bbox_from_cache(shapes_list, bbox_cache)

    if region_bbox is None:
        return _group_shapes_by_visual_rows_simple(shapes_list, row_threshold_points)

    # 执行 XY-Cut 分区
    regions = _xy_cut_partition(shapes_list, region_bbox, 0, bbox_cache, row_threshold_points)

    # 将 regions 转换为 rows
    return _xy_cut_regions_to_rows(regions, row_threshold_points, slide_size=slide_size)


def is_list_block(shape):
    """检测是否为列表块（参考 pptx2md 的逻辑）"""
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
        _debug_exc("is_list_block: 读取段落缩进失败", e)
        return False


def _get_single_line_plain_text(shape):
    """尝试从shape提取“可合并为一行”的纯文本（不做Markdown转义）。

    仅用于“编号shape + 标题shape”这种跨shape拼接的场景：
    - 必须是文本框
    - 仅1个段落
    - 不是列表块
    - 不是表格/图片等
    """
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

        if is_list_block(shape):
            return None

        text = tr.Paragraphs(1, 1).Text.strip()
        return text if text else None
    except Exception as e:
        _debug_exc("_get_single_line_plain_text: 提取单行文本失败", e)
        return None


def _render_texts_from_shape_row(row_shapes, skip_first_para_by_shape_id=None):
    """将同一视觉行内的多个shape渲染成Markdown行列表（不含末尾\\n）。

    关键：把“数字编号shape”和其右侧标题shape合并为Markdown有序列表项。
    """
    if not row_shapes:
        return []

    merged_lines = []
    used = set()
    skip_first_para_by_shape_id = skip_first_para_by_shape_id or {}

    # 尝试合并：找到第一个“纯数字”的shape，再找其右侧第一个普通文本shape
    num_i = None
    num_val = None
    for i, shape in enumerate(row_shapes):
        sid = _safe_shape_id(shape)
        if sid is not None and sid in skip_first_para_by_shape_id:
            continue
        t = _get_single_line_plain_text(shape)
        if t is None:
            continue
        m = re.fullmatch(r"(\d+)\.?", t)
        if m:
            num_i = i
            num_val = m.group(1)
            break

    title_i = None
    title_text = None
    if num_i is not None:
        for j in range(num_i + 1, len(row_shapes)):
            sid = _safe_shape_id(row_shapes[j])
            if sid is not None and sid in skip_first_para_by_shape_id:
                continue
            t = _get_single_line_plain_text(row_shapes[j])
            if t is None:
                continue
            if re.fullmatch(r"(\d+)\.?", t):
                continue
            title_i = j
            title_text = t
            break

    if num_i is not None and title_i is not None:
        merged_lines.append(f"{int(num_val)}. {_escape_md_text_line(title_text)}")
        used.add(num_i)
        used.add(title_i)

    for i, shape in enumerate(row_shapes):
        if i in used:
            continue
        sid = _safe_shape_id(shape)
        skip_text = skip_first_para_by_shape_id.get(sid) if sid is not None else None
        for text in extract_text_from_shape(shape, skip_first_para_text=skip_text):
            merged_lines.append(text)

    return merged_lines


def extract_text_from_shape(shape, skip_first_para_text=None):
    """从单个Shape中提取文本，支持列表层级格式"""
    texts = []

    # 检测图片 (msoPicture=13, msoLinkedPicture=11)
    try:
        if shape.Type in (13, 11):
            alt_text = ""
            try:
                alt_text = shape.AlternativeText
            except Exception as e:
                _debug_exc("extract_text_from_shape: 读取图片AlternativeText失败", e)
            if alt_text:
                safe_alt = str(alt_text).replace("\r", "").replace("\n", " ").replace("]", "\\]")
                texts.append(f"![图片: {safe_alt}]")
            else:
                texts.append("![图片]")
            return texts
    except Exception as e:
        _debug_exc("extract_text_from_shape: 读取shape.Type失败", e)

    if shape.HasTextFrame:
        try:
            tr = shape.TextFrame.TextRange
            para_count = tr.Paragraphs().Count

            if para_count == 0:
                return texts

            is_list = is_list_block(shape)
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
                        texts.append(_escape_md_text_line(text))
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
                        texts.append(f"{indent}{n}. {_escape_md_text_line(text)}")
                    else:
                        # 无序列表：保持旧行为
                        ol_counters.clear()
                        marker = "*" if is_list else "-"
                        texts.append(f"{indent}{marker} {_escape_md_text_line(text)}")
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

                    normalized = _strip_bullet_like_prefix(text)
                    if normalized is not None:
                        texts.append(f"{indent}- {_escape_md_text_line(normalized)}")
                        continue

                    texts.append(_escape_md_text_line(text))
        except Exception as e:
            _debug_exc("extract_text_from_shape: 解析TextFrame失败，尝试回退", e)
            # 回退到原始方式
            try:
                text = shape.TextFrame.TextRange.Text
                if text and text.strip():
                    texts.append(_escape_md_text_line(text))
            except Exception as e:
                _debug_exc("extract_text_from_shape: 回退读取TextRange.Text失败", e)

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
                        row.append(_escape_md_table_cell(cell))
                    except Exception as e:
                        _debug_exc("extract_text_from_shape: 读取表格单元格失败", e)
                        row.append("")
                rows.append(row)

            if rows:
                # Markdown表格
                col_count = len(rows[0]) if rows[0] else 0
                if col_count == 0:
                    return texts

                if TABLE_HEADER_MODE == "empty":
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
        _debug_exc("extract_text_from_shape: 处理表格失败", e)

    return texts


def extract_embedded_ppt(shape, ppt_app, activate_fn=None, loc_parts=None, depth=0, max_depth=5, ui_enabled=True):
    """提取嵌入的PPT内容（支持多层嵌套，路径以引用块形式输出，如：`路径：S2/E1/S1`）。

    Args:
        shape: OLE嵌入对象Shape（PowerPoint）
        ppt_app: PowerPoint.Application
        activate_fn: （可选）用于顶层嵌入时激活UI焦点的回调
        loc_parts: 当前位置路径分段，如 ["S2", "E1"]
        depth: 当前嵌套深度（0表示主PPT页内的嵌入PPT）
        max_depth: 最大递归深度（包含当前层），防止极端嵌套导致卡死
    """
    content = []
    embedded_ppt = None
    need_close = False
    allow_ui = bool(ui_enabled) and (int(depth) <= 0)  # 仅顶层嵌入允许走UI激活/DoVerb，避免多层状态污染
    loc_parts = list(loc_parts) if loc_parts else []
    try:
        prog_id = shape.OLEFormat.ProgID
        if 'PowerPoint' not in prog_id:
            prefix = _format_loc(loc_parts)
            if prefix:
                content.append(f"> [{prefix} 嵌入对象: {prog_id}]\n")
            else:
                content.append(f"> [嵌入对象: {prog_id}]\n")
            return content

        def _get_embedded_ppt():
            obj = shape.OLEFormat.Object
            _ = obj.Slides.Count
            return obj

        # 优先不走UI激活，直接尝试读取嵌入对象，减少前台/焦点依赖
        try:
            embedded_ppt = _wait_com(_get_embedded_ppt, 1, "extract_embedded_ppt: 读取嵌入PPT对象(未激活)")
        except Exception as e:
            _debug_exc("extract_embedded_ppt: 未激活读取失败", e)
            if not allow_ui:
                prefix = _format_loc(loc_parts)
                if prefix:
                    content.append(f"> [{prefix} 嵌入PPT无法直接读取，已跳过: {_format_exc(e, limit=80)}]\n")
                else:
                    content.append(f"> [嵌入PPT无法直接读取，已跳过: {_format_exc(e, limit=80)}]\n")
                return content

            _debug_exc("extract_embedded_ppt: 尝试激活/打开嵌入对象", e)
            if activate_fn is not None:
                _try_call(activate_fn, "extract_embedded_ppt: 激活嵌入对象失败")
            need_close = True
            _try_call(lambda: shape.OLEFormat.DoVerb(0), "extract_embedded_ppt: DoVerb(0)失败")
            embedded_ppt = _wait_com(_get_embedded_ppt, COM_EMBED_TIMEOUT_SEC,
                                    "extract_embedded_ppt: 读取嵌入PPT对象(激活后)")

        slide_count = _wait_com(lambda: embedded_ppt.Slides.Count, COM_EMBED_TIMEOUT_SEC,
                                "extract_embedded_ppt: 读取嵌入PPT幻灯片数量失败")

        # 获取嵌入 PPT 的幻灯片尺寸用于 XY-Cut
        embed_slide_size = None
        try:
            embed_width = float(embedded_ppt.PageSetup.SlideWidth)
            embed_height = float(embedded_ppt.PageSetup.SlideHeight)
            embed_slide_size = (embed_width, embed_height)
        except Exception as e:
            _debug_exc("extract_embedded_ppt: 读取嵌入PPT幻灯片尺寸失败", e)

        for i in range(1, slide_count + 1):
            eslide = _wait_com(lambda: embedded_ppt.Slides(i), COM_EMBED_TIMEOUT_SEC,
                               f"extract_embedded_ppt: 访问嵌入幻灯片{i}失败")

            slide_loc = loc_parts + [f"S{i}"]
            slide_h_level = TITLE_HEADING_LEVEL
            slide_loc_str = _format_loc(slide_loc)
            title_info = detect_slide_title(eslide)
            title_text = title_info.get("title") or f"嵌入幻灯片 {i}"
            content.append(_md_slide_heading_with_ref(slide_h_level, title_text, "嵌入幻灯片", i, slide_loc_str))
            slide_has_content = False
            skip_map = {}
            if title_info.get("shape_id") and title_info.get("title"):
                skip_map[int(title_info["shape_id"])] = str(title_info["title"])

            # 构建排除标题 shape 的 id 集合
            exclude_ids = set()
            if title_info.get("shape_id") and title_info.get("title"):
                exclude_ids.add(int(title_info["shape_id"]))

            # 标题shape可能同时包含副标题/正文（同一文本框多段），此处补回除标题外的段落
            for text in _extract_title_shape_extra_lines(eslide, title_info):
                content.append(f"{text}\n")
                if str(text).strip():
                    slide_has_content = True

            # 按视觉位置分行排序（并支持"编号shape + 标题shape"合并）
            shape_rows = []
            try:
                shape_rows = group_shapes_by_visual_rows(
                    list(eslide.Shapes),
                    ROW_THRESHOLD_POINTS,
                    enable_xy_cut=True,
                    slide_size=embed_slide_size,
                    exclude_shape_ids=exclude_ids,
                )
            except Exception as e:
                _debug_exc("extract_embedded_ppt: 枚举/排序嵌入幻灯片Shapes失败", e)

            embedded_shapes = []
            embedded_in_slide = 0

            # 第一轮：普通内容（跳过嵌入PPT对象，稍后递归处理）
            for row_shapes in shape_rows:
                normal_shapes = []
                for s in row_shapes:
                    try:
                        if s.Type == 7:
                            pid = s.OLEFormat.ProgID
                            if 'PowerPoint' in pid:
                                embedded_shapes.append(s)
                                continue
                            else:
                                content.append(f"> [{_format_loc(slide_loc)} 嵌入对象: {pid}]\n\n")
                                continue
                    except Exception as e:
                        _debug_exc("extract_embedded_ppt: 检测嵌入对象失败", e)
                    normal_shapes.append(s)

                for text in _render_texts_from_shape_row(normal_shapes, skip_first_para_by_shape_id=skip_map):
                    content.append(f"{text}\n")
                    if str(text).strip():
                        slide_has_content = True

            # 第二轮：递归处理嵌入PPT
            if int(depth) + 1 < int(max_depth):
                for s in embedded_shapes:
                    embedded_in_slide += 1
                    child_loc = slide_loc + [f"E{embedded_in_slide}"]
                    if slide_has_content:
                        content.append(_md_hr())
                    content.append(_md_embedded_ppt_marker(f"嵌入PPT #{embedded_in_slide}", _format_loc(child_loc)))
                    slide_has_content = True
                    try:
                        content.extend(extract_embedded_ppt(
                            s,
                            ppt_app,
                            activate_fn=None,
                            loc_parts=child_loc,
                            depth=int(depth) + 1,
                            max_depth=max_depth,
                            ui_enabled=ui_enabled,
                        ))
                    except Exception as e:
                        _debug_exc("extract_embedded_ppt: 递归提取嵌入PPT失败", e)
                        content.append(f"> [{_format_loc(child_loc)} 嵌入PPT提取失败: {_format_exc(e, limit=80)}]\n\n")
            elif embedded_shapes:
                content.append(f"> [{_format_loc(slide_loc)} 已达到最大嵌套深度{max_depth}，跳过更深嵌入]\n\n")

            # 在嵌入幻灯片之间添加分隔符（最后一个除外）
            if i < slide_count:
                content.append("\n---\n\n")

    except Exception as e:
        _debug_exc("extract_embedded_ppt: 提取嵌入PPT失败", e)
        content.append(f"> [提取失败: {str(e)[:80]}]\n")
    finally:
        if need_close:
            _close_embedded_object(ppt_app, embedded_ppt)

    return content


def extract_ppt_content(ppt_path, output_path=None, debug=False, ui=True):
    """提取PPT内容并保存为Markdown文件"""
    global DEBUG, TABLE_HEADER_MODE
    DEBUG = bool(debug)
    ui = bool(ui)
    ppt_path = os.path.abspath(ppt_path)

    if not os.path.exists(ppt_path):
        print(f"文件不存在: {ppt_path}")
        return False

    if output_path is None:
        # 默认输出到PPT所在目录，文件名与PPT一致
        ppt_dir = os.path.dirname(ppt_path)
        base_name = os.path.splitext(os.path.basename(ppt_path))[0]
        output_path = os.path.join(ppt_dir, base_name + ".md")
        output_path = get_unique_output_path(output_path)

    print(f"输入: {ppt_path}")
    print(f"输出: {output_path}")

    ppt_app = win32com.client.Dispatch("PowerPoint.Application")
    # PowerPoint在某些版本/策略下不允许将Application.Visible设置为False，
    # 且部分环境下 Presentations.Open(..., WithWindow=False) 也可能失败。
    # 因此 --no-ui 采取“尽量不干扰前台”的策略：禁用UI激活回退，同时尝试隐藏/最小化窗口；
    # 若无法隐藏，则回退可见但尽量保持最小化，确保流程可运行。
    if ui:
        _try_call(lambda: setattr(ppt_app, "Visible", True), "extract_ppt_content: 设置Visible=True失败")
    else:
        try:
            ppt_app.Visible = False
        except Exception as e:
            _debug_exc("extract_ppt_content: 设置Visible=False失败，回退Visible=True", e)
            _try_call(lambda: setattr(ppt_app, "Visible", True), "extract_ppt_content: 回退Visible=True失败")
    _try_call(lambda: setattr(ppt_app, "DisplayAlerts", 0), "extract_ppt_content: 设置DisplayAlerts失败")

    md = []
    presentation = None

    try:
        print("\n正在处理...")
        if ui:
            presentation = ppt_app.Presentations.Open(ppt_path, WithWindow=True)
        else:
            # 尽量无窗口打开；若失败则回退为有窗口打开（不同PowerPoint版本/策略差异较大）。
            try:
                presentation = ppt_app.Presentations.Open(ppt_path, WithWindow=False)
            except Exception as e:
                _debug_exc("extract_ppt_content: WithWindow=False打开失败，回退WithWindow=True", e)
                presentation = ppt_app.Presentations.Open(ppt_path, WithWindow=True)
        if not ui:
            # 尽量最小化窗口，避免抢焦点（不保证所有版本都有效）
            _try_call(lambda: setattr(ppt_app, "WindowState", 2), "extract_ppt_content: 最小化Application窗口失败")
            _try_call(lambda: setattr(ppt_app.ActiveWindow, "WindowState", 2), "extract_ppt_content: 最小化ActiveWindow失败")

        ppt_title = os.path.splitext(os.path.basename(ppt_path))[0]
        md.append(f"# {ppt_title}\n\n")

        total = _wait_com(lambda: presentation.Slides.Count, COM_OPEN_TIMEOUT_SEC, "extract_ppt_content: 等待PPT打开")
        embedded_count = 0

        # 获取幻灯片尺寸用于 XY-Cut
        slide_size = None
        try:
            slide_width = float(presentation.PageSetup.SlideWidth)
            slide_height = float(presentation.PageSetup.SlideHeight)
            slide_size = (slide_width, slide_height)
        except Exception as e:
            _debug_exc("extract_ppt_content: 读取幻灯片尺寸失败", e)

        for idx in range(1, total + 1):
            try:
                slide = presentation.Slides(idx)
            except Exception as e:
                print(f" 跳过(错误)")
                md.append(_md_slide_heading_with_ref(TITLE_HEADING_LEVEL, f"幻灯片 {idx}", "幻灯片", idx, f"S{idx}"))
                md.append(f"> [幻灯片访问错误]\n")
                md.append("\n---\n\n")
                continue

            print(f"  [{idx}/{total}]", end="")

            title_info = detect_slide_title(slide)
            title_text = title_info.get("title") or f"幻灯片 {idx}"
            md.append(_md_slide_heading_with_ref(TITLE_HEADING_LEVEL, title_text, "幻灯片", idx, f"S{idx}"))
            slide_has_content = False
            skip_map = {}
            if title_info.get("shape_id") and title_info.get("title"):
                skip_map[int(title_info["shape_id"])] = str(title_info["title"])

            # 构建排除标题 shape 的 id 集合
            exclude_ids = set()
            if title_info.get("shape_id") and title_info.get("title"):
                exclude_ids.add(int(title_info["shape_id"]))

            # 标题shape可能同时包含副标题/正文（同一文本框多段），此处补回除标题外的段落
            for text in _extract_title_shape_extra_lines(slide, title_info):
                md.append(f"{text}\n")
                if str(text).strip():
                    slide_has_content = True

            embedded_in_slide = 0
            embedded_shapes = []  # 收集嵌入PPT的Shape，稍后处理

            try:
                shape_rows = group_shapes_by_visual_rows(
                    list(slide.Shapes),
                    ROW_THRESHOLD_POINTS,
                    enable_xy_cut=True,
                    slide_size=slide_size,
                    exclude_shape_ids=exclude_ids,
                )
            except Exception as e:
                md.append(f"> [幻灯片读取错误: {str(e)[:50]}]\n")
                shape_rows = []

            # 第一轮：只处理普通内容（非嵌入PPT）
            for row_shapes in shape_rows:
                normal_shapes = []
                for shape in row_shapes:
                    # 检查是否为嵌入PPT，如果是则跳过，稍后处理
                    try:
                        if shape.Type == 7:
                            prog_id = shape.OLEFormat.ProgID
                            if 'PowerPoint' in prog_id:
                                embedded_shapes.append(shape)
                                continue
                            else:
                                md.append(f"> [嵌入对象: {prog_id}]\n\n")
                                continue
                    except Exception as e:
                        _debug_exc("extract_ppt_content: 检测嵌入对象失败", e)

                    normal_shapes.append(shape)

                for text in _render_texts_from_shape_row(normal_shapes, skip_first_para_by_shape_id=skip_map):
                    md.append(f"{text}\n")
                    if str(text).strip():
                        slide_has_content = True

            # 第二轮：处理嵌入PPT
            for shape in embedded_shapes:
                embedded_in_slide += 1
                embedded_count += 1
                if slide_has_content:
                    md.append(_md_hr())
                md.append(_md_embedded_ppt_marker(f"嵌入PPT #{embedded_in_slide}", f"S{idx}/E{embedded_in_slide}"))
                slide_has_content = True
                try:
                    def _activate():
                        win = presentation.Windows(1)
                        win.ViewType = 1
                        win.View.GotoSlide(idx)
                        shape.Select()
                        return True

                    activate_fn = None
                    if ui:
                        activate_fn = lambda: _wait_com(
                            _activate,
                            COM_EMBED_TIMEOUT_SEC,
                            f"extract_ppt_content: 激活嵌入对象失败(幻灯片{idx})"
                        )

                    md.extend(extract_embedded_ppt(
                        shape,
                        ppt_app,
                        activate_fn=activate_fn,
                        loc_parts=[f"S{idx}", f"E{embedded_in_slide}"],
                        depth=0,
                        max_depth=5,
                        ui_enabled=ui,
                    ))
                    slide_has_content = True
                except Exception as e:
                    _debug_exc("extract_ppt_content: 提取嵌入PPT失败", e)
                    md.append(f"> [嵌入PPT提取失败: {_format_exc(e, limit=80)}]\n\n")

            md.append("\n---\n\n")

            if embedded_in_slide:
                print(f" +{embedded_in_slide}个嵌入")
            else:
                print()

        # 先保存文件
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(''.join(md))

        print(f"\n完成: {total}张幻灯片, {embedded_count}个嵌入PPT")
        print(f"保存: {output_path}")

        return True

    except Exception as e:
        print(f"\n错误: {e}")
        traceback.print_exc()
        return False

    finally:
        if presentation is not None:
            try:
                presentation.Close()
            except Exception as e:
                _debug_exc("extract_ppt_content: presentation.Close失败", e)
        try:
            ppt_app.Quit()
        except Exception as e:
            _debug_exc("extract_ppt_content: ppt_app.Quit失败", e)


def main():
    parser = argparse.ArgumentParser(description='提取PPT内容为Markdown')
    parser.add_argument('input', help='PPT文件路径')
    parser.add_argument('-o', '--output', help='输出路径')
    parser.add_argument('--debug', action='store_true', help='输出调试日志与异常堆栈到stderr')
    parser.add_argument('--no-ui', action='store_true',
                        help='无UI模式：PowerPoint后台打开（WithWindow=False, Visible=False）。注意：若嵌入PPT无法直接读取，将跳过UI激活回退。')
    parser.add_argument('--table-header', choices=['first-row', 'empty'], default='first-row',
                        help='表格表头模式：first-row(默认，首行作表头) / empty(空表头，所有行按数据输出)')
    args = parser.parse_args()
    global TABLE_HEADER_MODE
    TABLE_HEADER_MODE = args.table_header
    success = extract_ppt_content(args.input, args.output, debug=args.debug, ui=(not args.no_ui))
    sys.exit(0 if success else 1)


if __name__ == '__main__':
    if len(sys.argv) == 1:
        print("用法: python extract_ppt.py <PPT文件> [-o 输出路径]")
        print("调试: python extract_ppt.py <PPT文件> --debug")
        print("表格: python extract_ppt.py <PPT文件> --table-header empty")
        print("默认输出: PPT所在目录，文件名与PPT一致（冲突时自动添加序号）")
    else:
        main()
