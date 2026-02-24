"""参数面板组件 - 主窗口右侧。"""

from pathlib import Path
import sys
from tkinter import filedialog
from typing import Dict, Any

import customtkinter as ctk
from CTkToolTip import CTkToolTip

from .. import theme


class ParamsPanel(ctk.CTkScrollableFrame):
    """右面板，包含所有转换参数设置。"""

    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self.configure(fg_color="transparent")

        # CustomTkinter 在 Windows 下默认 `yscrollincrement=1`（像素），而滚轮滚动使用
        # `-int(event.delta/6)` 个 "units"（通常每次 20），导致实际滚动距离偏小。
        # 这里仅针对本面板提高滚动增量，让滚轮更跟手。
        if sys.platform.startswith("win") and hasattr(self, "_parent_canvas"):
            try:
                # 20 units * 4px = 80px / tick，接近常见应用的手感。
                self._parent_canvas.configure(yscrollincrement=4)
            except Exception:
                # 兜底：避免因 customtkinter 内部实现变更导致 GUI 初始化失败。
                pass

        self._tooltips: list = []
        self._setup_ui()
        self._setup_tooltips()

    def _setup_ui(self):
        # 输出设置分组
        self._create_output_settings()

        # 图片选项分组
        self._create_image_options()

        # 内容处理分组
        self._create_content_options()

        # 高级选项分组
        self._create_advanced_options()

        # PPT 转换设置 (实验性)
        self._create_ppt_options()

    def _tip(self, widget, text: str):
        """为控件附加工具提示并保持引用。"""
        self._tooltips.append(CTkToolTip(widget, message=text, delay=0.25))

    def _create_group_frame(
        self,
        title: str,
        *,
        badge_text: str | None = None,
        badge_tip: str | None = None,
    ) -> ctk.CTkFrame:
        """创建带标题的分组框架。"""
        frame = ctk.CTkFrame(
            self,
            fg_color=theme.CARD_BG,
            corner_radius=8,
            border_width=1,
            border_color=theme.BORDER_COLOR,
        )
        frame.pack(fill="x", padx=10, pady=8)

        header = ctk.CTkFrame(frame, fg_color="transparent")
        header.pack(fill="x", padx=12, pady=(10, 5))

        label = ctk.CTkLabel(
            header,
            text=title,
            font=ctk.CTkFont(size=13, weight="bold"),
            anchor="w",
            text_color=theme.TEXT_PRIMARY,
        )
        label.pack(side="left")

        if badge_tip:
            self._tip(label, badge_tip)

        if badge_text:
            badge = ctk.CTkLabel(
                header,
                text=badge_text,
                font=ctk.CTkFont(size=13, weight="bold"),
                text_color=theme.BADGE_TEXT,
            )
            badge.pack(side="left", padx=(6, 0))

            if badge_tip:
                self._tip(badge, badge_tip)

        content = ctk.CTkFrame(frame, fg_color="transparent")
        content.pack(fill="x", padx=12, pady=(0, 12))

        return content

    def _create_output_settings(self):
        """创建输出设置分组。"""
        content = self._create_group_frame(
            "输出设置",
            badge_text="⁕",
            badge_tip="⁕ 表示本组为必填项",
        )

        # 输出目录
        dir_frame = ctk.CTkFrame(content, fg_color="transparent")
        dir_frame.pack(fill="x", pady=3)

        ctk.CTkLabel(dir_frame, text="输出目录:", width=80, anchor="w").pack(side="left")
        self.output_dir_var = ctk.StringVar()
        self.output_dir_entry = ctk.CTkEntry(
            dir_frame, textvariable=self.output_dir_var, width=200
        )
        self.output_dir_entry.pack(side="left", fill="x", expand=True, padx=5)
        ctk.CTkButton(
            dir_frame,
            text="浏览",
            width=60,
            fg_color=theme.BTN_NEUTRAL_BG,
            hover_color=theme.BTN_NEUTRAL_HOVER,
            text_color=theme.BTN_NEUTRAL_TEXT,
            command=self._browse_output_dir,
        ).pack(side="left")

        # 文件命名
        name_frame = ctk.CTkFrame(content, fg_color="transparent")
        name_frame.pack(fill="x", pady=3)

        self.naming_label = ctk.CTkLabel(name_frame, text="文件命名:", width=80, anchor="w")
        self.naming_label.pack(side="left")
        self.naming_var = ctk.StringVar(value="same")
        self.naming_segment = ctk.CTkSegmentedButton(
            name_frame,
            values=["与源文件同名", "自定义前缀"],
            variable=self.naming_var,
            command=self._on_naming_changed,
        )
        self.naming_segment.pack(side="left", padx=5)

        self.prefix_entry = ctk.CTkEntry(name_frame, width=100, placeholder_text="前缀")
        self.prefix_entry.pack(side="left", padx=5)
        self.prefix_entry.configure(state="disabled")

        # 输出格式
        format_frame = ctk.CTkFrame(content, fg_color="transparent")
        format_frame.pack(fill="x", pady=3)

        self.format_label = ctk.CTkLabel(format_frame, text="输出格式:", width=80, anchor="w")
        self.format_label.pack(side="left")
        self.format_var = ctk.StringVar(value="Markdown")
        self.format_segment = ctk.CTkSegmentedButton(
            format_frame,
            values=["Markdown", "Wiki", "Madoko", "Quarto"],
            variable=self.format_var,
        )
        self.format_segment.pack(side="left", padx=5)

    def _create_image_options(self):
        """创建图片选项分组。"""
        content = self._create_group_frame("图片选项")

        # 复选框行
        check_frame = ctk.CTkFrame(content, fg_color="transparent")
        check_frame.pack(fill="x", pady=3)

        self.disable_image_var = ctk.BooleanVar()
        self.disable_image_cb = ctk.CTkCheckBox(
            check_frame,
            text="禁用图片提取",
            variable=self.disable_image_var,
            command=self._on_disable_image_changed,
        )
        self.disable_image_cb.pack(side="left", padx=(0, 20))

        self.disable_wmf_var = ctk.BooleanVar()
        self.disable_wmf_cb = ctk.CTkCheckBox(
            check_frame, text="禁用 WMF 转换", variable=self.disable_wmf_var
        )
        self.disable_wmf_cb.pack(side="left")

        # 图片目录
        img_dir_frame = ctk.CTkFrame(content, fg_color="transparent")
        img_dir_frame.pack(fill="x", pady=3)

        self.image_dir_label = ctk.CTkLabel(img_dir_frame, text="图片目录:", width=80, anchor="w")
        self.image_dir_label.pack(side="left")
        self.image_dir_var = ctk.StringVar()
        self.image_dir_entry = ctk.CTkEntry(
            img_dir_frame, textvariable=self.image_dir_var, width=200
        )
        self.image_dir_entry.pack(side="left", fill="x", expand=True, padx=5)
        self.image_dir_btn = ctk.CTkButton(
            img_dir_frame,
            text="浏览",
            width=60,
            fg_color=theme.BTN_NEUTRAL_BG,
            hover_color=theme.BTN_NEUTRAL_HOVER,
            text_color=theme.BTN_NEUTRAL_TEXT,
            command=self._browse_image_dir,
        )
        self.image_dir_btn.pack(side="left")

        # 图片宽度
        width_frame = ctk.CTkFrame(content, fg_color="transparent")
        width_frame.pack(fill="x", pady=3)

        self.image_width_label = ctk.CTkLabel(width_frame, text="图片宽度:", width=80, anchor="w")
        self.image_width_label.pack(side="left")
        self.image_width_var = ctk.StringVar()
        self.image_width_entry = ctk.CTkEntry(
            width_frame,
            textvariable=self.image_width_var,
            width=100,
            placeholder_text="留空则不限制",
        )
        self.image_width_entry.pack(side="left", padx=5)
        ctk.CTkLabel(width_frame, text="px").pack(side="left")

    def _create_content_options(self):
        """创建内容处理选项分组。"""
        content = self._create_group_frame("内容处理")
        
        # 使用 Grid 布局实现两列对齐
        content.grid_columnconfigure(0, weight=1)
        content.grid_columnconfigure(1, weight=1)

        # Row 0
        self.enable_color_var = ctk.BooleanVar(value=True)
        self.enable_color_cb = ctk.CTkCheckBox(
            content, text="保留颜色标签", variable=self.enable_color_var
        )
        self.enable_color_cb.grid(row=0, column=0, sticky="w", padx=0, pady=5)

        self.enable_escaping_var = ctk.BooleanVar(value=True)
        self.enable_escaping_cb = ctk.CTkCheckBox(
            content, text="转义 Markdown 特殊字符", variable=self.enable_escaping_var
        )
        self.enable_escaping_cb.grid(row=0, column=1, sticky="w", padx=0, pady=5)

        # Row 1
        self.enable_notes_var = ctk.BooleanVar(value=True)
        self.enable_notes_cb = ctk.CTkCheckBox(
            content, text="提取演讲备注", variable=self.enable_notes_var
        )
        self.enable_notes_cb.grid(row=1, column=0, sticky="w", padx=0, pady=5)

        self.enable_slide_number_var = ctk.BooleanVar(value=True)
        self.enable_slide_number_cb = ctk.CTkCheckBox(
            content, text="添加幻灯片编号注释", variable=self.enable_slide_number_var
        )
        self.enable_slide_number_cb.grid(row=1, column=1, sticky="w", padx=0, pady=5)

        # Row 2
        self.enable_slides_var = ctk.BooleanVar()
        self.enable_slides_cb = ctk.CTkCheckBox(
            content, text="添加幻灯片分隔符", variable=self.enable_slides_var
        )
        self.enable_slides_cb.grid(row=2, column=0, sticky="w", padx=0, pady=5)

        self.compress_blank_lines_var = ctk.BooleanVar(value=True)
        self.compress_blank_lines_cb = ctk.CTkCheckBox(
            content,
            text="压缩连续空行",
            variable=self.compress_blank_lines_var,
        )
        self.compress_blank_lines_cb.grid(row=2, column=1, sticky="w", padx=0, pady=5)

        # Row 3
        self.keep_similar_titles_var = ctk.BooleanVar()
        self.keep_similar_titles_cb = ctk.CTkCheckBox(
            content,
            text="保留相似标题",
            variable=self.keep_similar_titles_var,
        )
        self.keep_similar_titles_cb.grid(row=3, column=0, sticky="w", padx=0, pady=5)

        # 最小文本块设置 (移动到右侧与上一行并列)
        block_frame = ctk.CTkFrame(content, fg_color="transparent")
        block_frame.grid(row=3, column=1, sticky="w", padx=0, pady=5)

        self.min_block_size_label = ctk.CTkLabel(block_frame, text="最小字符数:", anchor="w")
        self.min_block_size_label.pack(side="left")
        self.min_block_size_var = ctk.StringVar(value="15")
        self.min_block_size_entry = ctk.CTkEntry(
            block_frame, textvariable=self.min_block_size_var, width=50
        )
        self.min_block_size_entry.pack(side="left", padx=5)

    def _create_advanced_options(self):
        """创建高级选项分组。"""
        content = self._create_group_frame("高级选项")

        # 多列布局检测
        self.try_multi_column_var = ctk.BooleanVar()
        self.try_multi_column_cb = ctk.CTkCheckBox(
            content,
            text="尝试多列布局检测（处理速度较慢）",
            variable=self.try_multi_column_var,
        )
        self.try_multi_column_cb.pack(fill="x", pady=3)

        # 页码
        page_frame = ctk.CTkFrame(content, fg_color="transparent")
        page_frame.pack(fill="x", pady=3)

        self.page_label = ctk.CTkLabel(page_frame, text="仅转换页码:", anchor="w")
        self.page_label.pack(side="left")
        self.page_var = ctk.StringVar()
        self.page_entry = ctk.CTkEntry(
            page_frame,
            textvariable=self.page_var,
            width=100,
            placeholder_text="留空则转换全部",
        )
        self.page_entry.pack(side="left", padx=5)

        # 标题文件
        title_frame = ctk.CTkFrame(content, fg_color="transparent")
        title_frame.pack(fill="x", pady=3)

        self.title_path_label = ctk.CTkLabel(title_frame, text="标题列表文件:", anchor="w")
        self.title_path_label.pack(side="left")
        self.title_path_var = ctk.StringVar()
        self.title_path_entry = ctk.CTkEntry(
            title_frame, textvariable=self.title_path_var, width=180
        )
        self.title_path_entry.pack(side="left", fill="x", expand=True, padx=5)
        self.title_path_btn = ctk.CTkButton(
            title_frame,
            text="浏览",
            width=60,
            fg_color=theme.BTN_NEUTRAL_BG,
            hover_color=theme.BTN_NEUTRAL_HOVER,
            text_color=theme.BTN_NEUTRAL_TEXT,
            command=self._browse_title_file,
        )
        self.title_path_btn.pack(side="left")

    def _create_ppt_options(self):
        """创建 PPT 转换设置分组（实验性）。"""
        group = self._create_group_frame("PPT 转换设置", badge_text="实验性",
                                         badge_tip="仅对 .ppt 文件生效，需要 Windows + PowerPoint")

        self._ppt_widgets = []  # 收集可交互控件，用于灰置

        # 调试日志
        self.ppt_debug_var = ctk.BooleanVar(value=False)
        self.ppt_debug_cb = ctk.CTkCheckBox(
            group, text="调试日志", variable=self.ppt_debug_var,
            font=ctk.CTkFont(size=13),
        )
        self.ppt_debug_cb.pack(anchor="w", padx=15, pady=(10, 5))
        self._ppt_widgets.append(self.ppt_debug_cb)
        self._tip(self.ppt_debug_cb, "启用后将输出 COM 调试日志到日志面板")

        # 隐藏 PPT 窗口
        self.ppt_no_ui_var = ctk.BooleanVar(value=False)
        self.ppt_no_ui_cb = ctk.CTkCheckBox(
            group, text="隐藏 PowerPoint 窗口", variable=self.ppt_no_ui_var,
            font=ctk.CTkFont(size=13),
        )
        self.ppt_no_ui_cb.pack(anchor="w", padx=15, pady=5)
        self._ppt_widgets.append(self.ppt_no_ui_cb)
        self._tip(self.ppt_no_ui_cb, "后台运行 PowerPoint（可能影响嵌入 PPT 提取）")

        # 表格标题模式
        table_frame = ctk.CTkFrame(group, fg_color="transparent")
        table_frame.pack(fill="x", padx=15, pady=(5, 10))
        ctk.CTkLabel(
            table_frame, text="表格标题模式", font=ctk.CTkFont(size=13),
        ).pack(side="left")
        self.ppt_table_header_var = ctk.StringVar(value="首行作为标题")
        self.ppt_table_header_combo = ctk.CTkComboBox(
            table_frame,
            values=["首行作为标题", "空标题行"],
            variable=self.ppt_table_header_var,
            width=140, height=28, state="readonly",
        )
        self.ppt_table_header_combo.pack(side="right")
        self._ppt_widgets.append(self.ppt_table_header_combo)
        self._tip(self.ppt_table_header_combo, "first-row: 首行作表头 / empty: 所有行作数据")

        # 初始灰置
        self.set_ppt_group_enabled(False)

    def set_ppt_group_enabled(self, enabled: bool):
        """启用/禁用 PPT 参数组的所有子控件。"""
        state = "normal" if enabled else "disabled"
        for widget in self._ppt_widgets:
            widget.configure(state=state)

    def _browse_output_dir(self):
        path = filedialog.askdirectory(title="选择输出目录")
        if path:
            self.output_dir_var.set(path)

    def _setup_tooltips(self):
        """为所有参数控件附加工具提示。"""
        # 输出设置
        self._tip(self.output_dir_entry, "转换后的文件保存到哪个文件夹")
        self._tip(self.naming_label, "选择输出文件名：保持原名或添加自定义前缀")
        self._tip(self.prefix_entry, "输出文件名前面会加上这段文字")
        self._tip(self.format_label, "选择转换的目标文档格式")

        # 图片选项
        self._tip(self.disable_image_cb, "勾选后跳过所有图片，只转换文字")
        self._tip(self.disable_wmf_cb,
                  "WMF 是旧版 Office 的矢量图格式，勾选后跳过此类图片")
        image_dir_tip = "留空则默认输出到“输出目录/img”（未设置输出目录时，默认在 PPTX 同级目录/img）"
        self._tip(self.image_dir_label, image_dir_tip)
        self._tip(self.image_dir_entry, image_dir_tip)
        self._tip(self.image_dir_btn, image_dir_tip)
        image_width_tip = "图片在文档中的显示宽度（像素），留空则不限制"
        self._tip(self.image_width_label, image_width_tip)
        self._tip(self.image_width_entry, image_width_tip)

        # 内容处理
        self._tip(self.enable_color_cb, "勾选后保留文字颜色，生成颜色标记")
        self._tip(self.enable_escaping_cb,
                  "勾选后转义 Markdown 特殊字符（如 *、#、_ 等）")
        self._tip(self.enable_notes_cb, "勾选后提取幻灯片下方的演讲者备注")
        self._tip(self.enable_slides_cb,
                  "在每页幻灯片之间插入水平分隔线（---）")
        self._tip(self.enable_slide_number_cb,
                  "勾选后在输出中标注幻灯片页码")
        self._tip(self.keep_similar_titles_cb,
                  "多页标题相同时保留每个，并添加 (cont.) 后缀")
        self._tip(self.compress_blank_lines_cb, "将连续多行空行压缩为 1 行空行，输出更紧凑")
        self._tip(
            self.min_block_size_label,
            "少于此字符数的文本块会被跳过，用于过滤页眉页脚等干扰",
        )
        self._tip(self.min_block_size_entry,
                  "少于此字符数的文本块会被跳过，用于过滤页眉页脚等干扰")

        # 高级选项
        self._tip(self.try_multi_column_cb,
                  "自动识别幻灯片中的多列排版并正确转换")
        page_tip = "只转换指定页码，如 1,3,5-10；留空则全部转换"
        self._tip(self.page_label, page_tip)
        self._tip(self.page_entry, page_tip)
        title_path_tip = "文本文件，每行一个关键词，匹配的文字会被识别为标题；用缩进表示层级"
        self._tip(self.title_path_label, title_path_tip)
        self._tip(self.title_path_entry, title_path_tip)

    def _browse_image_dir(self):
        path = filedialog.askdirectory(title="选择图片目录")
        if path:
            self.image_dir_var.set(path)

    def _browse_title_file(self):
        path = filedialog.askopenfilename(
            title="选择标题列表文件",
            filetypes=[("文本文件", "*.txt"), ("所有文件", "*.*")],
        )
        if path:
            self.title_path_var.set(path)

    def _on_naming_changed(self, value):
        if value == "自定义前缀":
            self.prefix_entry.configure(state="normal")
        else:
            self.prefix_entry.configure(state="disabled")

    def _on_disable_image_changed(self):
        state = "disabled" if self.disable_image_var.get() else "normal"
        self.disable_wmf_cb.configure(state=state)
        self.image_dir_entry.configure(state=state)
        self.image_dir_btn.configure(state=state)
        self.image_width_entry.configure(state=state)

    def get_params(self) -> Dict[str, Any]:
        """收集所有参数值。"""
        return {
            "output_dir": self.output_dir_var.get(),
            "naming": "same" if "同名" in self.naming_segment.get() else "prefix",
            "prefix": self.prefix_entry.get(),
            "output_format": self.format_var.get().lower(),
            "disable_image": self.disable_image_var.get(),
            "disable_wmf": self.disable_wmf_var.get(),
            "image_dir": self.image_dir_var.get(),
            "image_width": self.image_width_var.get(),
            "enable_color": self.enable_color_var.get(),
            "enable_escaping": self.enable_escaping_var.get(),
            "enable_notes": self.enable_notes_var.get(),
            "enable_slides": self.enable_slides_var.get(),
            "enable_slide_number": self.enable_slide_number_var.get(),
            "keep_similar_titles": self.keep_similar_titles_var.get(),
            "compress_blank_lines": self.compress_blank_lines_var.get(),
            "min_block_size": self.min_block_size_var.get(),
            "try_multi_column": self.try_multi_column_var.get(),
            "page": self.page_var.get(),
            "title_path": self.title_path_var.get(),
            # PPT 转换选项
            "ppt_debug": self.ppt_debug_var.get(),
            "ppt_ui": not self.ppt_no_ui_var.get(),
            "ppt_table_header": "first-row" if "首行" in self.ppt_table_header_var.get() else "empty",
        }

    def set_params(self, params: Dict[str, Any]):
        """设置参数值（用于加载预设）。"""
        if "output_dir" in params:
            self.output_dir_var.set(params["output_dir"])
        if "naming" in params:
            # 避免硬编码中文 label：直接从控件取 values，再按索引设置。
            values = list(self.naming_segment.cget("values"))
            if len(values) >= 2:
                selected = values[1] if params["naming"] == "prefix" else values[0]
                self.naming_segment.set(selected)
                self._on_naming_changed(selected)
        if "prefix" in params:
            self.prefix_entry.delete(0, "end")
            self.prefix_entry.insert(0, params["prefix"] or "")
        if "output_format" in params:
            fmt_map = {"markdown": "Markdown", "wiki": "Wiki", "madoko": "Madoko", "quarto": "Quarto"}
            self.format_var.set(fmt_map.get(params["output_format"], "Markdown"))
        if "disable_image" in params:
            self.disable_image_var.set(params["disable_image"])
            self._on_disable_image_changed()
        if "disable_wmf" in params:
            self.disable_wmf_var.set(params["disable_wmf"])
        if "image_dir" in params:
            self.image_dir_var.set(params["image_dir"])
        if "image_width" in params:
            self.image_width_var.set(str(params["image_width"]) if params["image_width"] else "")
        if "enable_color" in params:
            self.enable_color_var.set(params["enable_color"])
        elif "disable_color" in params:
            self.enable_color_var.set(not params["disable_color"])
        if "enable_escaping" in params:
            self.enable_escaping_var.set(params["enable_escaping"])
        elif "disable_escaping" in params:
            self.enable_escaping_var.set(not params["disable_escaping"])
        if "enable_notes" in params:
            self.enable_notes_var.set(params["enable_notes"])
        elif "disable_notes" in params:
            self.enable_notes_var.set(not params["disable_notes"])
        if "enable_slides" in params:
            self.enable_slides_var.set(params["enable_slides"])
        if "enable_slide_number" in params:
            self.enable_slide_number_var.set(params["enable_slide_number"])
        elif "disable_slide_number" in params:
            self.enable_slide_number_var.set(not params["disable_slide_number"])
        if "keep_similar_titles" in params:
            self.keep_similar_titles_var.set(params["keep_similar_titles"])
        if "compress_blank_lines" in params:
            self.compress_blank_lines_var.set(params["compress_blank_lines"])
        if "min_block_size" in params:
            self.min_block_size_var.set(str(params["min_block_size"]))
        if "try_multi_column" in params:
            self.try_multi_column_var.set(params["try_multi_column"])
        if "ppt_debug" in params:
            self.ppt_debug_var.set(params["ppt_debug"])
        if "ppt_ui" in params:
            self.ppt_no_ui_var.set(not params["ppt_ui"])
        if "ppt_table_header" in params:
            th_map = {"first-row": "首行作为标题", "empty": "空标题行"}
            self.ppt_table_header_var.set(th_map.get(params["ppt_table_header"], "首行作为标题"))

    def reset_to_defaults(self):
        """将所有参数重置为默认值。"""
        self.output_dir_var.set("")
        self.format_var.set("Markdown")
        self.disable_image_var.set(False)
        self.disable_wmf_var.set(False)
        self.image_dir_var.set("")
        self.image_width_var.set("")
        self.enable_color_var.set(True)
        self.enable_escaping_var.set(True)
        self.enable_notes_var.set(True)
        self.enable_slides_var.set(False)
        self.enable_slide_number_var.set(True)
        self.keep_similar_titles_var.set(False)
        self.compress_blank_lines_var.set(True)
        self.min_block_size_var.set("15")
        self.try_multi_column_var.set(False)
        self.page_var.set("")
        self.title_path_var.set("")
        self._on_disable_image_changed()
        # PPT 转换选项
        self.ppt_debug_var.set(False)
        self.ppt_no_ui_var.set(False)
        self.ppt_table_header_var.set("首行作为标题")
