# -*- coding: utf-8 -*-
"""
周检查安全检查问题整改清单 + 整改回复 docx 生成脚本
按 PDF 样表格式，从 7月份整改回复单 xlsx 提取数据生成 docx
"""

import openpyxl
import os
import re
from docx import Document
from docx.shared import Pt, Cm, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ROW_HEIGHT_RULE
from docx.oxml.ns import qn

# 字号映射
SIZE_ERHAO = 22   # 二号
SIZE_SANHAO = 16  # 三号
SIZE_SIHAO = 14   # 四号
SIZE_WUHAO = 10.5 # 五号
ROW_HEIGHT_CM = Pt(50)  # 签章表三行行高 50pt（已固化，不得更改）

# ===== 配置 =====
XLSX_DIR = r'D:\Users\19586\Desktop\01_工程项目\7月份整改回复单'
OUTPUT_DIR = XLSX_DIR  # 输出到同一目录

# 项目信息
PROJECT_NAME = '梧桐苑项目'
INSPECTOR_ORG = '中铁置业梧桐苑项目检查组'
CONTRACTOR = '陕建总包'

# 签章人员
JL_PERSON = '陈军锋'  # 监理单位
SG_PERSON = '庞小强'  # 施工单位（项目经理）

# ===== 工具函数 =====

def iso_to_chinese(iso_date):
    """2026-07-24 -> 2026年7月24日"""
    m = re.match(r'(\d{4})-(\d{2})-(\d{2})', iso_date.strip())
    if not m:
        return iso_date
    y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
    return '%d年%d月%d日' % (y, mo, d)

def add_date(base_iso, delta_days):
    """日期加减"""
    from datetime import datetime, timedelta
    dt = datetime.strptime(base_iso.strip(), '%Y-%m-%d')
    dt += timedelta(days=delta_days)
    return dt.strftime('%Y-%m-%d')

def set_cell_font(cell, text, font_size=SIZE_SIHAO, bold=False, font_name='宋体'):
    """设置单元格文字和字体"""
    cell.text = ''
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(text)
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.name = font_name
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)

def set_run_font(run, font_size=SIZE_SIHAO, bold=False, font_name='宋体'):
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.name = font_name
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)

def add_paragraph(doc, text, font_size=SIZE_SIHAO, bold=False, alignment=WD_ALIGN_PARAGRAPH.LEFT, font_name='宋体', space_after=6, first_line_indent=None):
    """添加段落"""
    p = doc.add_paragraph()
    p.alignment = alignment
    if first_line_indent:
        p.paragraph_format.first_line_indent = Cm(first_line_indent)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.line_spacing = Pt(font_size * 2)  # 行距 = 字号×2
    run = p.add_run(text)
    set_run_font(run, font_size, bold, font_name)
    return p

def add_signature_table(doc, inspect_date_cn, person_jl=None, person_sg=None):
    """添加三列签章表格（行高50pt，文字四号，姓名留空手签）"""
    table = doc.add_table(rows=3, cols=3)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = 'Table Grid'

    # 设置行高
    for row in table.rows:
        row.height = ROW_HEIGHT_CM
        row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY

    # 第一行：单位名称
    set_cell_font(table.cell(0, 0), '建设单位（签章）', SIZE_SIHAO, False)
    set_cell_font(table.cell(0, 1), '监理单位（签章）', SIZE_SIHAO, False)
    set_cell_font(table.cell(0, 2), '施工单位（签章）', SIZE_SIHAO, False)

    # 第二行：签章人
    set_cell_font(table.cell(1, 0), '', SIZE_SIHAO, False)
    set_cell_font(table.cell(1, 1), person_jl or '', SIZE_SIHAO, False)
    set_cell_font(table.cell(1, 2), person_sg or '', SIZE_SIHAO, False)

    # 第三行：日期
    set_cell_font(table.cell(2, 0), inspect_date_cn, SIZE_SIHAO, False)
    set_cell_font(table.cell(2, 1), inspect_date_cn, SIZE_SIHAO, False)
    set_cell_font(table.cell(2, 2), inspect_date_cn, SIZE_SIHAO, False)

    return table


def generate_docx(xlsx_path, output_path):
    """从 xlsx 生成 docx"""
    wb = openpyxl.load_workbook(xlsx_path)
    ws = wb['表单']

    # 提取检查日期 - 用正则提取 YYYY-MM-DD
    a4 = ws['A4'].value or ''
    m = re.search(r'(\d{4}-\d{2}-\d{2})', a4)
    if m:
        inspect_iso = m.group(1)
    else:
        print('[ERROR] Cannot extract date from A4: %s' % a4)
        return

    reply_iso = add_date(inspect_iso, 1)
    inspect_cn = iso_to_chinese(inspect_iso)
    reply_cn = iso_to_chinese(reply_iso)

    # 提取问题
    problems = []
    for row in range(6, 14):
        v = ws.cell(row=row, column=2).value
        if v and '隐患内容' in str(v):
            content = str(v).strip()
            if content == '隐患内容：  \n\n整改要求：\n\n整改回复：':
                continue

            hazard = ''
            requirement = ''
            reply = ''
            for line in content.split('\n'):
                line = line.strip()
                if line.startswith('隐患内容'):
                    hazard = line.replace('隐患内容：', '').strip()
                elif line.startswith('整改要求'):
                    requirement = line.replace('整改要求：', '').strip()
                elif line.startswith('整改回复'):
                    reply = line.replace('整改回复：', '').strip()

            if hazard:
                problems.append({
                    'hazard': hazard,
                    'requirement': requirement,
                    'reply': reply or '已整改'
                })

    if not problems:
        print('[WARN] No problems found in %s' % xlsx_path)
        return

    # ===== 生成 docx =====
    doc = Document()

    # 设置页面
    for section in doc.sections:
        section.top_margin = Cm(2.5)
        section.bottom_margin = Cm(2.5)
        section.left_margin = Cm(3.0)
        section.right_margin = Cm(3.0)

    # ===== 第1页：周检查安全检查问题整改清单 =====
    # 标题（二号）
    add_paragraph(doc, '周检查安全检查问题整改清单', font_size=SIZE_ERHAO, bold=True,
                  alignment=WD_ALIGN_PARAGRAPH.CENTER, font_name='黑体', space_after=12)

    # 项目名称（三号）
    add_paragraph(doc, '项目名称：%s' % PROJECT_NAME, font_size=SIZE_SANHAO, space_after=12)

    # 签章表格（姓名留空，手签字）
    add_signature_table(doc, inspect_cn, '', '')

    # 空行
    add_paragraph(doc, '', font_size=10.5, space_after=6)

    # 问题描述
    add_paragraph(doc, '%s，%s检查提出的问题和隐患如下：' % (inspect_cn, INSPECTOR_ORG),
                  font_size=SIZE_SIHAO, space_after=6)

    # 问题列表
    for i, p in enumerate(problems, 1):
        add_paragraph(doc, '%d、%s；' % (i, p['hazard']), font_size=SIZE_SIHAO,
                      space_after=4, first_line_indent=0.74)

    # 分页
    doc.add_page_break()

    # ===== 第2页：整改回复 =====
    add_paragraph(doc, '整改回复', font_size=SIZE_ERHAO, bold=True,
                  alignment=WD_ALIGN_PARAGRAPH.CENTER, font_name='黑体', space_after=12)

    # 整改说明
    add_paragraph(doc, '%s，我单位（%s）对%s检查提出的问题和隐患，'
                  '项目部已进行整改，现将整改情况反馈如下：' % (reply_cn, CONTRACTOR, INSPECTOR_ORG),
                  font_size=SIZE_SIHAO, space_after=6, first_line_indent=0.74)

    # 整改列表
    for i, p in enumerate(problems, 1):
        # 整改措施 = 整改要求（已落实）
        measure = p['requirement'] if p['requirement'] else '已整改'
        add_paragraph(doc, '%d、%s（%s）；' % (i, p['hazard'], measure),
                      font_size=SIZE_SIHAO, space_after=4, first_line_indent=0.74)

    # 空行
    add_paragraph(doc, '', font_size=10.5, space_after=6)

    # 签章表格（姓名留空，手签字）
    add_signature_table(doc, reply_cn, '', '')

    # 保存
    doc.save(output_path)
    print('[OK] Generated: %s' % output_path)


def main():
    import sys
    # 支持命令行指定目录，默认 7月份
    if len(sys.argv) > 1:
        base_dir = sys.argv[1]
    else:
        base_dir = XLSX_DIR

    # 自动发现 xlsx 并按检查日期排序生成 docx
    items = []
    for fname in sorted(os.listdir(base_dir)):
        if not fname.endswith('.xlsx'):
            continue
        fpath = os.path.join(base_dir, fname)
        wb = openpyxl.load_workbook(fpath)
        ws = wb['表单']
        a4 = ws['A4'].value or ''
        m = re.search(r'(\d{4}-\d{2}-\d{2})', a4)
        if not m:
            print('[SKIP] No date in %s' % fname)
            continue
        inspect_iso = m.group(1)
        docx_name = '周检查安全检查问题整改清单_%s.docx' % inspect_iso.replace('-', '')
        items.append((inspect_iso, fpath, docx_name))

    # 按检查日期升序
    items.sort(key=lambda x: x[0])

    for inspect_iso, xlsx_path, docx_name in items:
        docx_path = os.path.join(base_dir, docx_name)
        generate_docx(xlsx_path, docx_path)


if __name__ == '__main__':
    main()
