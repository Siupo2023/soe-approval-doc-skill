#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
国企请示文档生成器
根据国有企业公文格式规范自动生成 Word 文档
"""

import os
import sys
from datetime import datetime
from docx import Document
from docx.shared import Pt, Inches, Mm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ============== 格式常量定义 ==============

# 页边距（毫米转换为 Twips，1mm ≈ 56.7 twips）
MARGIN_TOP_MM = 25
MARGIN_BOTTOM_MM = 25
MARGIN_LEFT_MM = 28
MARGIN_RIGHT_MM = 26

# 字号（磅）
TITLE_FONT_SIZE = 22  # 二号字
BODY_FONT_SIZE = 16   # 三号字

# 行距（磅）
TITLE_LINE_SPACING = 30  # 标题行距
BODY_LINE_SPACING = 28  # 正文行距

# 字体名称（按优先级）
TITLE_FONTS = ['方正小标宋简体', '方正小标宋', '华文中宋', '黑体', 'SimHei']
BODY_FONTS = ['仿宋GB2312', '仿宋', '仿宋_GB2312', 'FangSong', 'SimSun']

# 段落缩进（字符，约等于 2 字符 * 12 磅/字符 ≈ 24 磅 ≈ 0.85 cm）
FIRST_LINE_INDENT_CHARS = 2


def set_margins(section, top_mm, bottom_mm, left_mm, right_mm):
    """设置页边距（单位：毫米）"""
    section.top_margin = Mm(top_mm)
    section.bottom_margin = Mm(bottom_mm)
    section.left_margin = Mm(left_mm)
    section.right_margin = Mm(right_mm)


def set_chinese_font(run, font_names, size_pt, bold=False):
    """设置中文字体（支持备选字体列表）"""
    run.font.size = Pt(size_pt)
    run.font.bold = bold

    # 尝试设置中文字体
    for font_name in font_names:
        try:
            run.font.name = font_name
            # 设置东亚字体（对中文很重要）
            run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
            break
        except:
            continue


def set_paragraph_format(paragraph, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
                        line_spacing_pt=None, first_line_indent_chars=None,
                        space_before=None, space_after=None):
    """设置段落格式"""
    # 对齐方式
    paragraph.alignment = alignment

    # 行距
    if line_spacing_pt:
        paragraph.paragraph_format.line_rule = WD_LINE_SPACING.EXACTLY
        paragraph.paragraph_format.line_spacing = Pt(line_spacing_pt)

    # 首行缩进
    if first_line_indent_chars:
        # 2 字符缩进（约等于 0.85 cm 或 24 磅）
        paragraph.paragraph_format.first_line_indent = Inches(0.33)  # 约 0.85 cm

    # 段前段后间距
    if space_before is not None:
        paragraph.paragraph_format.space_before = Pt(space_before)
    if space_after is not None:
        paragraph.paragraph_format.space_after = Pt(space_after)


def add_title(doc, text):
    """添加标题（方正小标宋简体，二号字，30磅行距，居中）"""
    paragraph = doc.add_paragraph()
    run = paragraph.add_run(text)
    set_chinese_font(run, TITLE_FONTS, TITLE_FONT_SIZE)
    set_paragraph_format(paragraph,
                       alignment=WD_ALIGN_PARAGRAPH.CENTER,
                       line_spacing_pt=TITLE_LINE_SPACING)
    return paragraph


def add_salutation(doc, text):
    """添加称谓（如"公司领导："）"""
    paragraph = doc.add_paragraph()
    run = paragraph.add_run(text)
    set_chinese_font(run, BODY_FONTS, BODY_FONT_SIZE)
    set_paragraph_format(paragraph,
                       alignment=WD_ALIGN_PARAGRAPH.LEFT,
                       line_spacing_pt=BODY_LINE_SPACING)
    return paragraph


def add_body_paragraph(doc, text):
    """添加正文段落（仿宋GB2312，三号字，28磅行距，首行缩进2字符）"""
    paragraph = doc.add_paragraph()
    run = paragraph.add_run(text)
    set_chinese_font(run, BODY_FONTS, BODY_FONT_SIZE)
    set_paragraph_format(paragraph,
                       alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
                       line_spacing_pt=BODY_LINE_SPACING,
                       first_line_indent_chars=FIRST_LINE_INDENT_CHARS)
    return paragraph


def add_section_heading(doc, level, text):
    """
    添加小标题

    level: 1=一、, 2=（一）, 3=1., 4=（1）
    """
    prefixes = ['一、', '（一）', '1.', '（1）']
    if level < 1 or level > 4:
        level = 1

    full_text = f"{prefixes[level-1]}{text}"

    paragraph = doc.add_paragraph()
    run = paragraph.add_run(full_text)
    set_chinese_font(run, BODY_FONTS, BODY_FONT_SIZE, bold=(level == 1))

    # 一级标题不缩进，其他级别缩进
    indent = FIRST_LINE_INDENT_CHARS if level > 1 else None

    set_paragraph_format(paragraph,
                       alignment=WD_ALIGN_PARAGRAPH.LEFT,
                       line_spacing_pt=BODY_LINE_SPACING,
                       first_line_indent_chars=indent)
    return paragraph


def add_request_items(doc, items):
    """添加请示事项（不缩进的列表）"""
    paragraph = doc.add_paragraph()
    run = paragraph.add_run("现将有关事项请示如下：")
    set_chinese_font(run, BODY_FONTS, BODY_FONT_SIZE)
    set_paragraph_format(paragraph,
                       alignment=WD_ALIGN_PARAGRAPH.LEFT,
                       line_spacing_pt=BODY_LINE_SPACING)

    for i, item in enumerate(items, 1):
        paragraph = doc.add_paragraph()
        run = paragraph.add_run(f"{i}、{item}")
        set_chinese_font(run, BODY_FONTS, BODY_FONT_SIZE)
        set_paragraph_format(paragraph,
                           alignment=WD_ALIGN_PARAGRAPH.LEFT,
                           line_spacing_pt=BODY_LINE_SPACING,
                           first_line_indent_chars=FIRST_LINE_INDENT_CHARS)

    # 结束语
    paragraph = doc.add_paragraph()
    run = paragraph.add_run("妥否，请批示。")
    set_chinese_font(run, BODY_FONTS, BODY_FONT_SIZE)
    set_paragraph_format(paragraph,
                       alignment=WD_ALIGN_PARAGRAPH.LEFT,
                       line_spacing_pt=BODY_LINE_SPACING,
                       first_line_indent_chars=FIRST_LINE_INDENT_CHARS)


def add_attachments(doc, attachments):
    """
    添加附件说明

    attachments: list of attachment names
    格式：正文下空 1 行，左空 2 字
    """
    # 添加空行
    doc.add_paragraph()

    paragraph = doc.add_paragraph()
    text = "附件：" + " ".join([f"{i+1}.{name}" for i, name in enumerate(attachments)])
    run = paragraph.add_run(text)
    set_chinese_font(run, BODY_FONTS, BODY_FONT_SIZE)

    # 左空 2 字（约等于 0.85 cm 或 24 磅）
    paragraph.paragraph_format.left_indent = Inches(0.33)
    set_paragraph_format(paragraph,
                       alignment=WD_ALIGN_PARAGRAPH.LEFT,
                       line_spacing_pt=BODY_LINE_SPACING)


def add_signature(doc, department, date_str=None):
    """
    添加落款

    department: 部门名称
    date_str: 日期字符串（格式："YYYY年MM月DD日"），默认为当天
    """
    if date_str is None:
        date_str = datetime.now().strftime("%Y年%m月%d日")

    # 添加一些空行
    for _ in range(2):
        doc.add_paragraph()

    # 部门名称
    paragraph = doc.add_paragraph()
    run = paragraph.add_run(department)
    set_chinese_font(run, BODY_FONTS, BODY_FONT_SIZE)
    set_paragraph_format(paragraph,
                       alignment=WD_ALIGN_PARAGRAPH.RIGHT,
                       line_spacing_pt=BODY_LINE_SPACING)

    # 日期
    paragraph = doc.add_paragraph()
    run = paragraph.add_run(date_str)
    set_chinese_font(run, BODY_FONTS, BODY_FONT_SIZE)
    set_paragraph_format(paragraph,
                       alignment=WD_ALIGN_PARAGRAPH.RIGHT,
                       line_spacing_pt=BODY_LINE_SPACING)


def generate_approval_doc(output_path, title, salutation, sections,
                         request_items, attachments, department):
    """
    生成完整的请示文档

    参数:
        output_path: 输出文件路径（.docx）
        title: 文档标题
        salutation: 称谓（如"公司领导："）
        sections: list of dict，每个 dict 包含:
            - level: 标题层级 (1-4)
            - heading: 标题文本
            - content: list of 段落文本
        request_items: list of 请示事项
        attachments: list of 附件名称
        department: 落款部门
    """
    doc = Document()

    # 设置页边距
    set_margins(doc.sections[0],
               MARGIN_TOP_MM,
               MARGIN_BOTTOM_MM,
               MARGIN_LEFT_MM,
               MARGIN_RIGHT_MM)

    # 标题
    add_title(doc, title)

    # 称谓
    add_salutation(doc, salutation)

    # 各个部分
    for section in sections:
        level = section.get('level', 1)
        heading = section.get('heading', '')
        content = section.get('content', [])

        # 小标题
        add_section_heading(doc, level, heading)

        # 正文段落
        for para_text in content:
            if para_text.strip():
                add_body_paragraph(doc, para_text)

    # 请示事项
    if request_items:
        add_request_items(doc, request_items)

    # 附件
    if attachments:
        add_attachments(doc, attachments)

    # 落款
    add_signature(doc, department)

    # 保存文档
    doc.save(output_path)
    print(f"✅ 文档已生成：{output_path}")


# ============== 命令行接口 ==============

def main():
    """命令行接口"""
    if len(sys.argv) < 2:
        print("用法: python3 generate_approval_doc.py <输出文件.docx>")
        print("或从 Claude 调用 generate_approval_doc() 函数")
        sys.exit(1)

    output_path = sys.argv[1]

    # 示例：生成一个测试文档
    generate_approval_doc(
        output_path=output_path,
        title="关于采购办公设备的请示",
        salutation="公司领导：",
        sections=[
            {
                'level': 1,
                'heading': '基本情况',
                'content': [
                    '我部门现有办公设备已使用超过5年，部分设备已出现老化现象，影响日常办公效率。',
                    '为保障部门正常运转，提高工作效率，现申请更新部分办公设备。'
                ]
            },
            {
                'level': 1,
                'heading': '必要性与可行性',
                'content': [
                    '必要性：现有设备故障频发，维修成本逐年增加，影响工作进度。',
                    '可行性：经市场调研，相关设备价格合理，预算在部门年度预算范围内。'
                ]
            },
            {
                'level': 1,
                'heading': '采购方案',
                'content': [
                    '拟采购电脑10台，打印机2台，投影仪1台。',
                    '预算总金额约15万元。'
                ]
            }
        ],
        request_items=[
            '同意采购办公设备，预算金额15万元',
            '同意按规定程序进行采购'
        ],
        attachments=['办公设备采购清单', '报价单'],
        department='综合管理部'
    )


if __name__ == '__main__':
    main()
