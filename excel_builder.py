#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
excel_builder.py
================
把 AI 生成的人群包 JSON，写成符合 old_skill 标准格式的 .xlsx 文件。

Excel 结构完全沿用 old_skill/references/meitu_audience_format.md 的定义：
    1. 标题行（合并 A:D，品牌主色深色底 + 白字）
    2. 数据说明（5 条）
    3. 人群包汇总（1 行公式）
    4. 基础人群定向（4 行：性别 / 年龄 / 兴趣 / 消费）
    5. 定制人群包圈选规则（核心，每个子标签 1 行）
    6. 备注（7-8 条）

列定义：A 人群包(20) / B 人群标签(26) / C 秀秀单标签逻辑之一(68) / D 人群包量级(14)

另外额外加一个「行业洞察与参考来源」工作表，放联网搜索到的洞察和链接，
方便策划同事写方案时引用（不影响第一个工作表的标准格式）。
"""

import os
import re
from datetime import datetime

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ============================================================
# 兜底配色（AI 没给品牌色时用美图品牌色）
# ============================================================
DEFAULT_COLORS = {
    "title": "8C0B28",    # 深红（标题行）
    "section": "F11D48",  # 美图红（板块标题）
    "header": "FF6B8A",   # 亮粉红（表头）
    "basic": "FFE4EC",    # 浅粉（基础定向底色）
}

# 人群包行背景浅色轮换池（AI 没给 color 时按顺序取）
FALLBACK_SEGMENT_COLORS = [
    "FFF3F6", "FFF7E6", "F1F8F4", "EEF4FD", "FBF0FB",
    "F5F5F5", "FFF9E8", "EEF9FB", "F6F1FB", "EDEDED",
]


def _safe_hex(value, fallback):
    """
    校验颜色值。
    openpyxl 只认 6 位（或 8 位）十六进制，AI 可能给成 "#F11D48" 或乱写，
    所以这里统一清洗：去掉 #、转大写、长度不对就用兜底色。
    """
    text = str(value or "").strip().lstrip("#").upper()
    if re.fullmatch(r"[0-9A-F]{6}", text):
        return text
    if re.fullmatch(r"[0-9A-F]{8}", text):   # 带透明度的，取后 6 位
        return text[-6:]
    return fallback


def _clean_sheet_title(name):
    """
    工作表名清洗：Excel 不允许 : \\ / ? * [ ] 这些字符，且长度上限 31。
    """
    text = str(name or "定制人群包").strip()
    text = re.sub(r"[:\\/?*\[\]]", "", text)
    text = text[:24] or "定制人群包"
    return "{}定制人群包".format(text)[:31]


def safe_filename(name):
    """把品牌名等用户输入清洗成可以当文件名的字符串。"""
    text = str(name or "").strip()
    text = re.sub(r"[\\/:*?\"<>|\r\n\t]", "_", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text[:60] or "未命名品牌"


def build_audience_excel(data, output_dir, form=None, extra_meta=None):
    """
    生成人群包 Excel 文件。

    参数：
        data       : AI 返回并解析好的人群包字典（见 skill_rules.OUTPUT_SCHEMA）
        output_dir : 输出文件夹（如 ./output）
        form       : 网页表单内容（用于补全品牌名等）
        extra_meta : 额外信息字典，如 {"web_search_used": True, "model": "..."}

    返回：生成好的 Excel 文件完整路径
    """
    data = data or {}
    form = form or {}
    extra_meta = extra_meta or {}

    # ---------- 基础信息 ----------
    brand_name = str(data.get("brand_name") or form.get("brand_name") or "未命名品牌").strip()
    brand_name_en = str(data.get("brand_name_en") or "").strip()
    product_name = str(form.get("product_name") or "").strip()

    colors = data.get("colors") or {}
    color_title = _safe_hex(colors.get("title"), DEFAULT_COLORS["title"])
    color_section = _safe_hex(colors.get("section"), DEFAULT_COLORS["section"])
    color_header = _safe_hex(colors.get("header"), DEFAULT_COLORS["header"])
    color_basic = _safe_hex(colors.get("basic"), DEFAULT_COLORS["basic"])

    # ---------- 新建工作簿 ----------
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = _clean_sheet_title(brand_name_en or brand_name)

    # ---------- 字体（按 old_skill 样式规范）----------
    title_font = Font(name="微软雅黑", size=16, bold=True, color="FFFFFF")
    section_font = Font(name="微软雅黑", size=12, bold=True, color="FFFFFF")
    header_font = Font(name="微软雅黑", size=10, bold=True, color="FFFFFF")
    content_font = Font(name="微软雅黑", size=10, color="333333")
    content_bold = Font(name="微软雅黑", size=10, bold=True, color="333333")
    note_font = Font(name="微软雅黑", size=9, color="666666")

    # ---------- 填充色 ----------
    title_fill = PatternFill("solid", start_color=color_title, end_color=color_title)
    section_fill = PatternFill("solid", start_color=color_section, end_color=color_section)
    header_fill = PatternFill("solid", start_color=color_header, end_color=color_header)
    basic_fill = PatternFill("solid", start_color=color_basic, end_color=color_basic)
    white_fill = PatternFill("solid", start_color="FFFFFF", end_color="FFFFFF")

    # ---------- 边框与对齐 ----------
    thin_border = Border(
        left=Side(style="thin", color="B0B0B0"),
        right=Side(style="thin", color="B0B0B0"),
        top=Side(style="thin", color="B0B0B0"),
        bottom=Side(style="thin", color="B0B0B0"),
    )
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)
    left_top_align = Alignment(horizontal="left", vertical="top", wrap_text=True)

    # ---------- 列宽 ----------
    for col, width in {"A": 20, "B": 26, "C": 68, "D": 14}.items():
        ws.column_dimensions[col].width = width

    row = 1

    # ============================================================
    # 板块 1：标题行
    # ============================================================
    title_text = "{} x 美图定制人群包圈选规则".format(brand_name)
    if product_name:
        title_text = "{} {} x 美图定制人群包圈选规则".format(brand_name, product_name)

    ws.merge_cells("A{r}:D{r}".format(r=row))
    cell = ws.cell(row=row, column=1, value=title_text)
    cell.font = title_font
    cell.fill = title_fill
    cell.alignment = center_align
    ws.row_dimensions[row].height = 40
    row += 2

    # ============================================================
    # 板块 2：数据说明（5 条）
    # ============================================================
    row = _write_section_title(ws, row, "数据说明：", section_font, section_fill, left_align)

    data_notes = _as_str_list(data.get("data_notes"))
    if not data_notes:
        data_notes = [
            "1. 整体逻辑通过品牌目标消费人群特征匹配美图人群画像；",
            "2. 涉及的APP：美图系场景（美图秀秀、美颜相机等）；",
            "3. 人群包采用近6个月中符合圈选规则的人群；",
            "4. 品牌背景信息：{}。".format(brand_name),
            "5. 本次核心场景：待补充。",
        ]
    for note in data_notes:
        ws.merge_cells("A{r}:D{r}".format(r=row))
        c = ws.cell(row=row, column=1, value=note)
        c.font = note_font
        c.alignment = left_align
        ws.row_dimensions[row].height = 20
        row += 1
    row += 1

    # ============================================================
    # 板块 3：人群包汇总（1 行公式）
    # ============================================================
    row = _write_section_title(ws, row, "人群包汇总", section_font, section_fill, left_align)

    summary_text = str(data.get("summary_text") or "").strip()
    if not summary_text:
        seg_names = [str(s.get("name")) for s in (data.get("segments") or []) if s.get("name")]
        summary_text = "【基础定向】∩【{}】".format(" ∪ ".join(seg_names) if seg_names else "各定制人群包 ∪ 种子人群包")

    ws.merge_cells("A{r}:D{r}".format(r=row))
    cell = ws.cell(row=row, column=1, value=summary_text)
    cell.font = content_font
    cell.fill = basic_fill
    cell.alignment = left_align
    ws.row_dimensions[row].height = 36
    row += 2

    # ============================================================
    # 板块 4：基础人群定向（4 行，A 列合并）
    # ============================================================
    row = _write_section_title(ws, row, "基础人群定向", section_font, section_fill, left_align)

    # 表头
    for col_idx, h in enumerate(["基础人群定向", "人群维度", "秀秀单标签逻辑", "人群包量级"], 1):
        c = ws.cell(row=row, column=col_idx, value=h)
        c.font = header_font
        c.fill = header_fill
        c.alignment = center_align
        c.border = thin_border
    ws.row_dimensions[row].height = 28
    row += 1

    basic_items = data.get("basic_targeting") or []
    if not basic_items:
        basic_items = [
            {"dimension": "性别维度", "logic": "女 / 男，整体男女全覆盖"},
            {"dimension": "年龄维度", "logic": "18-40岁核心人群"},
            {"dimension": "兴趣爱好维度", "logic": "安装应用偏好 + 社区兴趣偏好"},
            {"dimension": "中高消费维度", "logic": "中等及以上消费力"},
        ]

    basic_start = row
    for i, item in enumerate(basic_items):
        if isinstance(item, dict):
            dimension = str(item.get("dimension") or "")
            logic = str(item.get("logic") or "")
            scale = str(item.get("scale") or "")
        else:
            # 容错：万一 AI 给的是列表 ["性别维度", "女/男"]
            parts = list(item) + ["", "", ""]
            dimension, logic, scale = str(parts[0]), str(parts[1]), str(parts[2])

        fill = basic_fill if i % 2 == 0 else white_fill
        for c_idx, val in enumerate(["", dimension, logic, scale], 1):
            c = ws.cell(row=row, column=c_idx, value=val)
            c.fill = fill
            c.font = content_bold if c_idx == 2 else content_font
            c.alignment = left_align if c_idx == 3 else center_align
            c.border = thin_border
        # 行高按内容长度估算（兴趣维度往往有两行）
        ws.row_dimensions[row].height = max(24, min(60, len(logic) // 3 + logic.count("\n") * 16 + 10))
        row += 1

    if row - 1 > basic_start:
        ws.merge_cells("A{s}:A{e}".format(s=basic_start, e=row - 1))
    merged = ws.cell(row=basic_start, column=1, value="基础人群定向")
    merged.font = content_bold
    merged.alignment = center_align
    merged.fill = basic_fill
    merged.border = thin_border
    row += 1

    # ============================================================
    # 板块 5：定制人群包圈选规则（核心）
    # ============================================================
    row = _write_section_title(
        ws, row, "{}定制人群包圈选规则".format(brand_name),
        section_font, section_fill, left_align,
    )

    for col_idx, h in enumerate(["人群包", "人群标签", "秀秀单标签逻辑之一", "人群包量级"], 1):
        c = ws.cell(row=row, column=col_idx, value=h)
        c.font = header_font
        c.fill = header_fill
        c.alignment = center_align
        c.border = thin_border
    ws.row_dimensions[row].height = 28
    row += 1

    segments = data.get("segments") or []
    tag_count = 0

    for seg_index, seg in enumerate(segments):
        if not isinstance(seg, dict):
            continue
        seg_name = str(seg.get("name") or "人群包{}".format(seg_index + 1))
        tags = seg.get("tags") or []
        if not tags:
            continue

        # 该人群包的行背景色：AI 给了就用，没给按轮换池取
        seg_color = _safe_hex(
            seg.get("color"),
            FALLBACK_SEGMENT_COLORS[seg_index % len(FALLBACK_SEGMENT_COLORS)],
        )
        seg_fill = PatternFill("solid", start_color=seg_color, end_color=seg_color)

        seg_start_row = row
        for tag_item in tags:
            if isinstance(tag_item, dict):
                tag = str(tag_item.get("tag") or "")
                logic = str(tag_item.get("logic") or "")
                scale = str(tag_item.get("scale") or "")
            else:
                parts = list(tag_item) + ["", "", ""]
                tag, logic, scale = str(parts[0]), str(parts[1]), str(parts[2])

            # A 列只在该人群包首行写名称，其余留空后面合并
            ws.cell(row=row, column=1, value=seg_name if row == seg_start_row else "")
            ws.cell(row=row, column=2, value=tag)
            ws.cell(row=row, column=3, value=logic)
            ws.cell(row=row, column=4, value=scale)

            for c_idx in range(1, 5):
                c = ws.cell(row=row, column=c_idx)
                c.font = content_bold if c_idx == 1 else content_font
                c.alignment = left_top_align if c_idx == 3 else center_align
                c.border = thin_border
                c.fill = seg_fill

            # 行高按圈选规则文字长度自动计算（沿用 old_skill 的算法）
            logic_lines = logic.count("\n") + 1
            ws.row_dimensions[row].height = max(36, min(90, len(logic) // 2 + logic_lines * 18))
            row += 1
            tag_count += 1

        # 合并该人群包的 A 列
        if row - 1 > seg_start_row:
            ws.merge_cells("A{s}:A{e}".format(s=seg_start_row, e=row - 1))
            top = ws.cell(row=seg_start_row, column=1)
            top.value = seg_name
            top.font = content_bold
            top.alignment = center_align
            top.fill = seg_fill

    row += 1

    # ============================================================
    # 板块 6：备注（7-8 条）
    # ============================================================
    row = _write_section_title(ws, row, "备注", section_font, section_fill, left_align)

    remarks = _as_str_list(data.get("remarks"))
    if not remarks:
        remarks = [
            "1. 以上人群包量级为预估量级，实际投放时以系统最终圈选结果为准；",
            "2. 各人群包之间采用「并集（∪）」逻辑组合，基础定向与定制人群包之间采用「交集（∩）」逻辑；",
            "3. 影像识别基于美图AI图像识别技术，识别用户上传/编辑图片中的品牌LOGO、场景、物品等元素；",
            "4. LBS场景识别基于用户打开APP时的地理位置信息，匹配线下消费场所POI数据；",
            "5. 原生需求指用户在美图秀秀产品中主动使用的功能/配方/滤镜等行为数据。",
        ]
    for remark in remarks:
        ws.merge_cells("A{r}:D{r}".format(r=row))
        c = ws.cell(row=row, column=1, value=remark)
        c.font = note_font
        c.alignment = left_align
        ws.row_dimensions[row].height = 20
        row += 1

    # 末尾加一行生成信息，方便追溯是谁什么时候生成的
    row += 1
    ws.merge_cells("A{r}:D{r}".format(r=row))
    gen_info = "本文件由「美图内部 · 广告人群包生成工具」于 {} 生成｜共 {} 个人群包 / {} 个子标签｜联网搜索：{}".format(
        datetime.now().strftime("%Y-%m-%d %H:%M"),
        len([s for s in segments if isinstance(s, dict) and s.get("tags")]),
        tag_count,
        "已开启" if extra_meta.get("web_search_used") else "未开启",
    )
    c = ws.cell(row=row, column=1, value=gen_info)
    c.font = Font(name="微软雅黑", size=8, color="999999")
    c.alignment = left_align

    # ============================================================
    # 附加工作表：行业洞察与参考来源
    # ============================================================
    _write_insight_sheet(wb, data, form, color_section, color_header)

    # ============================================================
    # 保存文件
    # ============================================================
    os.makedirs(output_dir, exist_ok=True)
    filename = "{brand}{product} x 美图定制人群包圈选规则_{ts}.xlsx".format(
        brand=safe_filename(brand_name),
        product=" " + safe_filename(product_name) if product_name else "",
        ts=datetime.now().strftime("%Y%m%d_%H%M%S"),
    )
    output_path = os.path.join(output_dir, filename)
    wb.save(output_path)
    return output_path


# ============================================================
# 内部小工具
# ============================================================

def _write_section_title(ws, row, text, font, fill, align):
    """写一个板块标题行（合并 A:D），返回下一行行号。"""
    ws.merge_cells("A{r}:D{r}".format(r=row))
    cell = ws.cell(row=row, column=1, value=text)
    cell.font = font
    cell.fill = fill
    cell.alignment = align
    ws.row_dimensions[row].height = 28
    return row + 1


def _as_str_list(value):
    """把 AI 返回的内容统一转成字符串列表，非列表/空值返回空列表。"""
    if not value:
        return []
    if isinstance(value, str):
        return [value]
    if isinstance(value, list):
        return [str(v).strip() for v in value if str(v).strip()]
    return []


def _write_insight_sheet(wb, data, form, color_section, color_header):
    """
    第二个工作表：把联网搜索到的行业洞察、竞品名单、参考链接和本次需求存档。
    """
    ws2 = wb.create_sheet("行业洞察与参考来源")
    ws2.column_dimensions["A"].width = 22
    ws2.column_dimensions["B"].width = 96

    section_font = Font(name="微软雅黑", size=12, bold=True, color="FFFFFF")
    section_fill = PatternFill("solid", start_color=color_section, end_color=color_section)
    label_font = Font(name="微软雅黑", size=10, bold=True, color="333333")
    body_font = Font(name="微软雅黑", size=10, color="333333")
    label_fill = PatternFill("solid", start_color=color_header, end_color=color_header)
    left_align = Alignment(horizontal="left", vertical="top", wrap_text=True)
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

    r = 1

    def section(title):
        """写板块标题，返回下一行。"""
        nonlocal_row = r
        ws2.merge_cells("A{x}:B{x}".format(x=nonlocal_row))
        c = ws2.cell(row=nonlocal_row, column=1, value=title)
        c.font = section_font
        c.fill = section_fill
        c.alignment = Alignment(horizontal="left", vertical="center")
        ws2.row_dimensions[nonlocal_row].height = 26
        return nonlocal_row + 1

    def kv(label, value, row_idx):
        """写一行"标签 - 内容"，返回下一行。"""
        lc = ws2.cell(row=row_idx, column=1, value=label)
        lc.font = label_font
        lc.fill = label_fill
        lc.alignment = center_align
        vc = ws2.cell(row=row_idx, column=2, value=str(value or "—"))
        vc.font = body_font
        vc.alignment = left_align
        text_len = len(str(value or ""))
        ws2.row_dimensions[row_idx].height = max(20, min(120, text_len // 2 + 18))
        return row_idx + 1

    # 一、本次需求存档
    r = section("一、本次需求输入（存档）")
    for label, key in [
        ("品牌名称", "brand_name"),
        ("广告产品/活动", "product_name"),
        ("广告目标", "ad_goal"),
        ("目标人群描述", "audience_desc"),
        ("合作内容/卖点", "cooperation"),
        ("竞品品牌", "competitors"),
        ("投放预算", "budget"),
        ("投放地区", "regions"),
        ("其他补充", "extra"),
    ]:
        r = kv(label, (form or {}).get(key), r)
    r += 1

    # 二、竞品名单（AI 补充的）
    r = section("二、竞品品牌名单（AI 补充，国际+国产）")
    competitors = _as_str_list(data.get("competitors"))
    r = kv("竞品名单", "、".join(competitors) if competitors else "—", r)
    r += 1

    # 三、行业洞察
    r = section("三、联网搜索得到的行业人群洞察 / 竞品投放策略 / 消费趋势")
    insights = _as_str_list(data.get("industry_insights"))
    if insights:
        for i, text in enumerate(insights, 1):
            r = kv("洞察 {}".format(i), text, r)
    else:
        r = kv("洞察", "本次未获取到联网搜索洞察内容", r)
    r += 1

    # 四、参考来源
    r = section("四、参考资料来源")
    sources = data.get("sources") or []
    if isinstance(sources, list) and sources:
        for i, src in enumerate(sources, 1):
            if isinstance(src, dict):
                title = src.get("title") or "参考资料 {}".format(i)
                url = src.get("url") or ""
                r = kv(str(title)[:20], url or "（无链接）", r)
            else:
                r = kv("来源 {}".format(i), str(src), r)
    else:
        r = kv("来源", "—", r)

    return ws2
