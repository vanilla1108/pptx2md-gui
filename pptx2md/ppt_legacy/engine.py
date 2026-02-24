# -*- coding: utf-8 -*-
"""PPT COM 转换入口。

⚠ 线程安全：本模块使用模块级 _log_cb / _progress_cb 引用，
  同一进程内同一时刻仅允许一个 extract_ppt_content() 调用。
  若未来需支持并行转换，需改为实例化或传参模式。
"""

import os
import sys
import time
import traceback
import re

from pptx2md.ppt_legacy.config import ExtractConfig, ConversionCancelled

from pptx2md.ppt_legacy.renderer_markdown import (
    format_loc as _format_loc,
    md_comment as _md_comment,
    md_embedded_ppt_marker as _md_embedded_ppt_marker,
    md_hr as _md_hr,
    md_slide_heading_with_ref as _md_slide_heading_with_ref,
)

from pptx2md.ppt_legacy.extractor_core import (
    build_image_extract_context_core as _build_image_extract_context_core,
    build_image_placeholder_markdown_core as _build_image_placeholder_markdown_core,
    build_title_render_context_core as _build_title_render_context_core,
    detect_slide_title_core as _detect_slide_title_core,
    escape_md_table_cell_core as _escape_md_table_cell_core,
    escape_md_text_line_core as _escape_md_text_line_core,
    extract_title_shape_extra_lines_core as _extract_title_shape_extra_lines_core,
    find_shape_by_id_in_slide_core as _find_shape_by_id_in_slide_core,
    first_paragraph_text_core as _first_paragraph_text_core,
    extract_text_from_shape_core as _extract_text_from_shape_core,
    export_shape_image_markdown_core as _export_shape_image_markdown_core,
    get_unique_output_path_core as _get_unique_output_path_core,
    get_single_line_plain_text_core as _get_single_line_plain_text_core,
    is_title_candidate_shape_core as _is_title_candidate_shape_core,
    is_list_block_core as _is_list_block_core,
    looks_like_brief_list_item_core as _looks_like_brief_list_item_core,
    next_export_image_path_core as _next_export_image_path_core,
    normalize_md_link_path_core as _normalize_md_link_path_core,
    process_shape_rows_core as _process_shape_rows_core,
    read_shape_alt_text_core as _read_shape_alt_text_core,
    render_shape_row_with_number_merge as _render_shape_row_with_number_merge,
    split_manual_ordered_prefix_core as _split_manual_ordered_prefix_core,
    safe_shape_id_core as _safe_shape_id_core,
    strip_bullet_like_prefix_core as _strip_bullet_like_prefix_core,
)

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

# ---------------------------------------------------------------------------
# 回调基础设施（由 extract_ppt_content 入口设置/清理）
# ---------------------------------------------------------------------------

_log_cb = None       # (level: str, message: str) -> None
_progress_cb = None  # (current: int, total: int, slide_name: str) -> None


def _log(level: str, msg: str):
    """统一日志出口。有回调用回调，无回调退化为 print。"""
    if _log_cb:
        _log_cb(level, f"[PPT] {msg}")
    else:
        target = sys.stderr if level in ("ERROR", "DEBUG") else sys.stdout
        print(msg, file=target)


def _debug(msg: str):
    """调试日志。仅 DEBUG=True 时输出。"""
    if not DEBUG:
        return
    _log("DEBUG", msg)


# ============================================================================
# 兼容封装层（对外函数名保持不变，内部委派到 extractor_core）
# ============================================================================

def _safe_shape_id(shape):
    return _safe_shape_id_core(shape)


def _first_paragraph_text(shape):
    """读取shape第一段文本（纯文本，strip后返回）。"""
    return _first_paragraph_text_core(shape, debug_exc_fn=_debug_exc)


def _is_title_candidate_shape(shape):
    """判断shape是否可能是"标题行"候选（不保证一定是标题）。"""
    return _is_title_candidate_shape_core(
        shape,
        is_list_block_fn=is_list_block,
        debug_exc_fn=_debug_exc,
    )


def detect_slide_title(slide):
    """识别单页幻灯片的第一行标题。

    优先使用 slide.Shapes.Title；否则从文本shape中按"靠上 + 字号大 + 文本短"择优。
    返回：
        {
          "title": str|None,
          "shape_id": int|None,   # 标题所在shape id（用于正文中跳过重复输出）
        }
    """
    return _detect_slide_title_core(
        slide,
        safe_shape_id_fn=_safe_shape_id,
        first_paragraph_text_fn=_first_paragraph_text,
        is_title_candidate_shape_fn=_is_title_candidate_shape,
    )


def _find_shape_by_id_in_slide(slide, shape_id):
    """在 slide.Shapes 中按 shape.Id 查找 Shape（失败则返回 None）。"""
    return _find_shape_by_id_in_slide_core(
        slide,
        shape_id,
        safe_shape_id_fn=_safe_shape_id,
        debug_exc_fn=_debug_exc,
    )


def _extract_title_shape_extra_lines(slide, title_info):
    """补回"标题shape里除第一段标题外的其他段落内容"。

    说明：
    - 当前实现会在布局分析中剔除 title_shape_id（避免影响 XY-Cut 多栏切割）
    - 但部分PPT把"标题 + 副标题/要点/正文"放在同一个文本框（同一个 shape）
      仅剔除会导致副标题/要点丢失（例如"本节提要"、课程目标列表等）
    - 因此这里把 title_shape 的"除第一段标题外"的内容补回到正文输出中
    """
    return _extract_title_shape_extra_lines_core(
        slide,
        title_info,
        find_shape_by_id_fn=_find_shape_by_id_in_slide,
        extract_text_from_shape_fn=extract_text_from_shape,
        debug_exc_fn=_debug_exc,
    )


def _format_exc(e, limit=120):
    msg = str(e) if e is not None else ""
    msg = msg.replace("\r", " ").replace("\n", " ")
    if len(msg) > limit:
        msg = msg[:limit] + "..."
    return f"{type(e).__name__}: {msg}" if msg else type(e).__name__


def _debug_exc(context, e):
    if not DEBUG:
        return
    _log("DEBUG", f"{context}: {_format_exc(e)}")
    try:
        tb = getattr(e, "__traceback__", None)
        if tb is not None and not _log_cb:
            # 仅无回调（CLI 独立运行）时打印完整堆栈到 stderr
            traceback.print_exception(type(e), e, tb, file=sys.stderr)
    except Exception:
        pass


def _wait_com(action, timeout_sec, context, cancel_event=None):
    """轮询等待COM调用成功；支持取消；超时则抛出最后一次异常。"""
    deadline = time.time() + timeout_sec
    start = time.time()
    attempts = 0
    last_exc = None
    first_exc = None
    while time.time() <= deadline:
        if cancel_event and cancel_event.is_set():
            raise ConversionCancelled()
        attempts += 1
        try:
            result = action()
            if DEBUG and attempts > 1:
                elapsed = time.time() - start
                _debug(f"{context}: 成功(重试{attempts - 1}次, {elapsed:.2f}s)")
            return result
        except ConversionCancelled:
            raise
        except Exception as e:
            if first_exc is None:
                first_exc = e
                _debug(f"{context}: 首次失败: {_format_exc(e)}")
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
    return _escape_md_text_line_core(text)


def _escape_md_table_cell(text):
    """转义表格单元格内容，避免破坏管道表格结构。"""
    return _escape_md_table_cell_core(text)


def _strip_bullet_like_prefix(text):
    """若文本以常见"项目符号样式字符"开头，则去掉该符号并返回剩余正文；否则返回None。

    仅用于普通段落（非脚本自身输出的引用/提示行）。
    """
    return _strip_bullet_like_prefix_core(text)


def _split_manual_ordered_prefix(text):
    """识别形如"1、内容"的手打编号，返回 (序号, 正文)。"""
    return _split_manual_ordered_prefix_core(text)


def _looks_like_brief_list_item(text, max_len=48):
    """粗略判断该文本是否更像列表项而非完整段落。"""
    return _looks_like_brief_list_item_core(text, max_len=max_len)


def get_unique_output_path(base_path):
    """获取唯一的输出路径，如果文件已存在则添加序号

    Args:
        base_path: 基础文件路径

    Returns:
        不冲突的文件路径
    """
    return _get_unique_output_path_core(base_path, path_exists_fn=os.path.exists)


def _normalize_md_link_path(path):
    """把Windows路径规范化为Markdown链接路径。"""
    return _normalize_md_link_path_core(path)


def _read_shape_alt_text(shape):
    """读取图片替代文本（失败时返回空串）。"""
    return _read_shape_alt_text_core(shape, debug_exc_fn=_debug_exc)


def _build_image_placeholder_markdown(shape=None, alt_text=None):
    """生成图片占位标注（禁用导图或导图失败时使用）。"""
    return _build_image_placeholder_markdown_core(
        shape=shape,
        alt_text=alt_text,
        read_shape_alt_text_fn=_read_shape_alt_text,
    )


def _build_image_extract_context(output_path, extract_images=True, image_dir=None):
    """构建图片导出上下文。"""
    return _build_image_extract_context_core(
        output_path,
        extract_images=extract_images,
        image_dir=image_dir,
    )


def _next_export_image_path(image_ctx, image_loc=None, shape=None):
    """按顺序生成唯一图片文件名。"""
    return _next_export_image_path_core(
        image_ctx,
        image_loc=image_loc,
        shape=shape,
        safe_shape_id_fn=_safe_shape_id,
    )


def _export_shape_image_markdown(shape, image_ctx=None, image_loc=None):
    """导出shape图片并返回Markdown图片语法；失败时回退占位标注。"""
    return _export_shape_image_markdown_core(
        shape,
        image_ctx=image_ctx,
        image_loc=image_loc,
        read_shape_alt_text_fn=_read_shape_alt_text,
        build_image_placeholder_markdown_fn=_build_image_placeholder_markdown,
        debug_exc_fn=_debug_exc,
        makedirs_fn=os.makedirs,
        next_export_image_path_fn=_next_export_image_path,
        wait_com_fn=_wait_com,
        com_open_timeout_sec=COM_OPEN_TIMEOUT_SEC,
        file_exists_fn=os.path.exists,
        relpath_fn=os.path.relpath,
        normalize_md_link_path_fn=_normalize_md_link_path,
    )


# ============================================================================
# 版式分析与排序（XY-Cut + 视觉行分组）
# ============================================================================

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
        _debug(f"自适应阈值: 中位数高度={median_height:.1f}磅, 阈值={threshold:.1f}磅 (样本数={len(heights)})")

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
    return _is_list_block_core(shape, debug_exc_fn=_debug_exc)


def _get_single_line_plain_text(shape):
    """尝试从shape提取"可合并为一行"的纯文本（不做Markdown转义）。

    仅用于"编号shape + 标题shape"这种跨shape拼接的场景：
    - 必须是文本框
    - 仅1个段落
    - 不是列表块
    - 不是表格/图片等
    """
    return _get_single_line_plain_text_core(
        shape,
        is_list_block_fn=is_list_block,
        debug_exc_fn=_debug_exc,
    )


def _render_texts_from_shape_row(row_shapes, skip_first_para_by_shape_id=None, image_ctx=None, loc_prefix=None):
    """将同一视觉行内的多个shape渲染成Markdown行列表（不含末尾\\n）。

    关键：把"数字编号shape"和其右侧标题shape合并为Markdown有序列表项。
    """
    return _render_shape_row_with_number_merge(
        row_shapes,
        skip_first_para_by_shape_id=skip_first_para_by_shape_id,
        image_ctx=image_ctx,
        loc_prefix=loc_prefix,
        safe_shape_id_fn=_safe_shape_id,
        get_single_line_plain_text_fn=_get_single_line_plain_text,
        escape_md_text_line_fn=_escape_md_text_line,
        extract_text_from_shape_fn=extract_text_from_shape,
    )


def extract_text_from_shape(shape, skip_first_para_text=None, image_ctx=None, image_loc=None):
    """从单个Shape中提取文本，支持列表层级格式"""
    return _extract_text_from_shape_core(
        shape,
        skip_first_para_text=skip_first_para_text,
        image_ctx=image_ctx,
        image_loc=image_loc,
        table_header_mode=TABLE_HEADER_MODE,
        export_shape_image_markdown_fn=_export_shape_image_markdown,
        debug_exc_fn=_debug_exc,
        is_list_block_fn=is_list_block,
        split_manual_ordered_prefix_fn=_split_manual_ordered_prefix,
        looks_like_brief_list_item_fn=_looks_like_brief_list_item,
        escape_md_text_line_fn=_escape_md_text_line,
        strip_bullet_like_prefix_fn=_strip_bullet_like_prefix,
        escape_md_table_cell_fn=_escape_md_table_cell,
    )


# ============================================================================
# COM 编排主流程（主PPT + 嵌入PPT）
# ============================================================================

def extract_embedded_ppt(shape, ppt_app, activate_fn=None, loc_parts=None, depth=0, max_depth=5,
                         ui_enabled=True, image_ctx=None, cancel_event=None):
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
                content.append(_md_comment(f"{prefix} embedded-object: {prog_id}"))
            else:
                content.append(_md_comment(f"embedded-object: {prog_id}"))
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
                    content.append(_md_comment(
                        f"{prefix} embedded-ppt-skip: {_format_exc(e, limit=80)}"
                    ))
                else:
                    content.append(_md_comment(f"embedded-ppt-skip: {_format_exc(e, limit=80)}"))
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
            if cancel_event and cancel_event.is_set():
                raise ConversionCancelled()
            eslide = _wait_com(lambda: embedded_ppt.Slides(i), COM_EMBED_TIMEOUT_SEC,
                               f"extract_embedded_ppt: 访问嵌入幻灯片{i}失败")

            slide_loc = loc_parts + [f"S{i}"]
            slide_h_level = TITLE_HEADING_LEVEL
            slide_loc_str = _format_loc(slide_loc)
            title_ctx = _build_title_render_context_core(
                eslide,
                fallback_title=f"嵌入幻灯片 {i}",
                detect_slide_title_fn=detect_slide_title,
                extract_title_shape_extra_lines_fn=_extract_title_shape_extra_lines,
            )
            title_text = title_ctx["title_text"]
            content.append(_md_slide_heading_with_ref(slide_h_level, title_text, "嵌入幻灯片", i, slide_loc_str))
            slide_has_content = False
            skip_map = title_ctx["skip_map"]
            exclude_ids = title_ctx["exclude_ids"]

            # 标题shape可能同时包含副标题/正文（同一文本框多段），此处补回除标题外的段落
            for text in title_ctx["extra_lines"]:
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
            row_lines, row_embedded_shapes = _process_shape_rows_core(
                shape_rows,
                slide_loc=_format_loc(slide_loc),
                row_renderer_fn=_render_texts_from_shape_row,
                skip_map=skip_map,
                image_ctx=image_ctx,
                embedded_object_line_fn=lambda pid: _md_comment(f"{_format_loc(slide_loc)} embedded-object: {pid}"),
                debug_exc_fn=_debug_exc,
                debug_context_prefix="extract_embedded_ppt",
            )
            embedded_shapes.extend(row_embedded_shapes)
            for text in row_lines:
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
                            image_ctx=image_ctx,
                            cancel_event=cancel_event,
                        ))
                    except ConversionCancelled:
                        raise
                    except Exception as e:
                        _debug_exc("extract_embedded_ppt: 递归提取嵌入PPT失败", e)
                        content.append(_md_comment(
                            f"{_format_loc(child_loc)} embedded-ppt-fail: {_format_exc(e, limit=80)}"
                        ) + "\n")
            elif embedded_shapes:
                content.append(_md_comment(
                    f"{_format_loc(slide_loc)} max-depth-reached: {max_depth}, skip deeper embedded"
                ) + "\n")

            # 在嵌入幻灯片之间添加分隔符（最后一个除外）
            if i < slide_count:
                content.append("\n---\n\n")

    except ConversionCancelled:
        raise
    except Exception as e:
        _debug_exc("extract_embedded_ppt: 提取嵌入PPT失败", e)
        content.append(_md_comment(f"extract-fail: {str(e)[:80]}"))
    finally:
        if need_close:
            _close_embedded_object(ppt_app, embedded_ppt)

    return content


def _apply_runtime_config(config):
    """应用运行时配置到旧全局开关，保持行为兼容。"""
    global DEBUG, TABLE_HEADER_MODE
    DEBUG = bool(config.debug)
    TABLE_HEADER_MODE = str(config.table_header)


def extract_ppt_content(config=None, log_callback=None, progress_callback=None, cancel_event=None,
                        # 兼容旧调用方式的位置参数（deprecated）
                        ppt_path=None, output_path=None, debug=False, ui=True,
                        extract_images=True, image_dir=None, table_header="first-row"):
    """提取PPT内容并保存为Markdown文件。

    参数:
        config: ExtractConfig 实例。
        log_callback: (level: str, message: str) -> None
        progress_callback: (current: int, total: int, slide_name: str) -> None
        cancel_event: threading.Event 或 None。

    返回:
        bool: 转换是否成功。
    """
    global _log_cb, _progress_cb
    _log_cb = log_callback
    _progress_cb = progress_callback
    try:
        return _extract_ppt_content_inner(config, cancel_event,
                                          ppt_path, output_path, debug, ui,
                                          extract_images, image_dir, table_header)
    finally:
        _log_cb = None
        _progress_cb = None


def _extract_ppt_content_inner(config, cancel_event,
                               ppt_path, output_path, debug, ui,
                               extract_images, image_dir, table_header):
    """extract_ppt_content 的实际执行体。"""
    import win32com.client

    if config is None:
        if isinstance(ppt_path, ExtractConfig):
            config = ppt_path
        else:
            config = ExtractConfig(
                input_path=ppt_path,
                output_path=output_path,
                debug=debug,
                ui=ui,
                extract_images=extract_images,
                image_dir=image_dir,
                table_header=table_header,
            )

    _apply_runtime_config(config)
    ui = bool(config.ui)
    ppt_path = os.path.abspath(config.input_path)
    output_path = config.output_path

    if not os.path.exists(ppt_path):
        _log("ERROR", f"文件不存在: {ppt_path}")
        return False

    if output_path is None:
        ppt_dir = os.path.dirname(ppt_path)
        base_name = os.path.splitext(os.path.basename(ppt_path))[0]
        output_path = os.path.join(ppt_dir, base_name + ".md")
        output_path = get_unique_output_path(output_path)

    image_ctx = _build_image_extract_context(output_path, extract_images=extract_images, image_dir=image_dir)
    if image_ctx.get("enabled"):
        try:
            os.makedirs(image_ctx["dir"], exist_ok=True)
        except Exception as e:
            _log("WARNING", f"图片目录创建失败，回退为占位标注: {image_ctx.get('dir')} ({_format_exc(e, limit=80)})")
            image_ctx["enabled"] = False

    _log("INFO", f"输入: {ppt_path}")
    _log("INFO", f"输出: {output_path}")
    if image_ctx.get("enabled"):
        _log("INFO", f"图片: 启用 (目录: {image_ctx.get('dir')})")
    else:
        _log("INFO", "图片: 禁用 (输出占位标注)")

    ppt_app = win32com.client.Dispatch("PowerPoint.Application")
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
        if cancel_event and cancel_event.is_set():
            raise ConversionCancelled()

        _log("INFO", "正在处理...")
        if ui:
            presentation = ppt_app.Presentations.Open(ppt_path, WithWindow=True)
        else:
            try:
                presentation = ppt_app.Presentations.Open(ppt_path, WithWindow=False)
            except Exception as e:
                _debug_exc("extract_ppt_content: WithWindow=False打开失败，回退WithWindow=True", e)
                presentation = ppt_app.Presentations.Open(ppt_path, WithWindow=True)
        if not ui:
            _try_call(lambda: setattr(ppt_app, "WindowState", 2), "extract_ppt_content: 最小化Application窗口失败")
            _try_call(lambda: setattr(ppt_app.ActiveWindow, "WindowState", 2), "extract_ppt_content: 最小化ActiveWindow失败")

        ppt_title = os.path.splitext(os.path.basename(ppt_path))[0]
        md.append(f"# {ppt_title}\n\n")

        total = _wait_com(lambda: presentation.Slides.Count, COM_OPEN_TIMEOUT_SEC,
                          "extract_ppt_content: 等待PPT打开", cancel_event=cancel_event)
        embedded_count = 0

        slide_size = None
        try:
            slide_width = float(presentation.PageSetup.SlideWidth)
            slide_height = float(presentation.PageSetup.SlideHeight)
            slide_size = (slide_width, slide_height)
        except Exception as e:
            _debug_exc("extract_ppt_content: 读取幻灯片尺寸失败", e)

        for idx in range(1, total + 1):
            if cancel_event and cancel_event.is_set():
                raise ConversionCancelled()

            try:
                slide = presentation.Slides(idx)
            except Exception as e:
                _log("WARNING", f"[{idx}/{total}] 跳过(错误)")
                md.append(_md_slide_heading_with_ref(TITLE_HEADING_LEVEL, f"幻灯片 {idx}", "幻灯片", idx, f"S{idx}"))
                md.append(_md_comment("slide-access-error"))
                md.append("\n---\n\n")
                continue

            title_ctx = _build_title_render_context_core(
                slide,
                fallback_title=f"幻灯片 {idx}",
                detect_slide_title_fn=detect_slide_title,
                extract_title_shape_extra_lines_fn=_extract_title_shape_extra_lines,
            )
            title_text = title_ctx["title_text"]

            _log("INFO", f"[{idx}/{total}] {title_text}")
            if _progress_cb:
                _progress_cb(idx, total, title_text)

            md.append(_md_slide_heading_with_ref(TITLE_HEADING_LEVEL, title_text, "幻灯片", idx, f"S{idx}"))
            slide_has_content = False
            skip_map = title_ctx["skip_map"]
            exclude_ids = title_ctx["exclude_ids"]

            for text in title_ctx["extra_lines"]:
                md.append(f"{text}\n")
                if str(text).strip():
                    slide_has_content = True

            embedded_in_slide = 0
            embedded_shapes = []

            try:
                shape_rows = group_shapes_by_visual_rows(
                    list(slide.Shapes),
                    ROW_THRESHOLD_POINTS,
                    enable_xy_cut=True,
                    slide_size=slide_size,
                    exclude_shape_ids=exclude_ids,
                )
            except Exception as e:
                md.append(_md_comment(f"slide-read-error: {str(e)[:50]}"))
                shape_rows = []

            row_lines, row_embedded_shapes = _process_shape_rows_core(
                shape_rows,
                slide_loc=f"S{idx}",
                row_renderer_fn=_render_texts_from_shape_row,
                skip_map=skip_map,
                image_ctx=image_ctx,
                embedded_object_line_fn=lambda prog_id: _md_comment(f"embedded-object: {prog_id}"),
                debug_exc_fn=_debug_exc,
                debug_context_prefix="extract_ppt_content",
            )
            embedded_shapes.extend(row_embedded_shapes)
            for text in row_lines:
                md.append(f"{text}\n")
                if str(text).strip():
                    slide_has_content = True

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
                            f"extract_ppt_content: 激活嵌入对象失败(幻灯片{idx})",
                            cancel_event=cancel_event,
                        )

                    md.extend(extract_embedded_ppt(
                        shape,
                        ppt_app,
                        activate_fn=activate_fn,
                        loc_parts=[f"S{idx}", f"E{embedded_in_slide}"],
                        depth=0,
                        max_depth=5,
                        ui_enabled=ui,
                        image_ctx=image_ctx,
                        cancel_event=cancel_event,
                    ))
                    slide_has_content = True
                except ConversionCancelled:
                    raise
                except Exception as e:
                    _debug_exc("extract_ppt_content: 提取嵌入PPT失败", e)
                    md.append(_md_comment(f"embedded-ppt-fail: {_format_exc(e, limit=80)}") + "\n")

            md.append("\n---\n\n")

            if embedded_in_slide:
                _log("INFO", f"  +{embedded_in_slide}个嵌入")

        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(''.join(md))

        _log("SUCCESS", f"完成: {total}张幻灯片, {embedded_count}个嵌入PPT")
        _log("INFO", f"保存: {output_path}")

        return True

    except ConversionCancelled:
        _log("WARNING", "转换已取消")
        return False
    except Exception as e:
        _log("ERROR", f"转换出错: {_format_exc(e)}")
        if DEBUG:
            traceback.print_exc(file=sys.stderr)
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
