# -*- coding: utf-8 -*-
r"""
工程管理部通知单生成器（固定模板模式，2026-09-07 用户定稿固化，**格式不可改变**）

用法：
  # 默认：以固定模板为底稿复制后替换内容（推荐，格式100%保真）
  python gen_engineering_notice.py --photo-dir <照片目录> --out <输出docx>

  # 从零重绘（仅在模板文件缺失时兜底）
  python gen_engineering_notice.py --rebuild --photo-dir <照片目录> --out <输出docx>

固定模板（不可改变）：
  assets/工程管理部通知单_固定模板_2026-09-07.docx
  （源自 K:\备份\Desktop_Backup\01_工程项目\鄠邑项目考核资料\工程管理部通知单2026-6-24.docx，
    经用户 2026-09-07 手工定稿确认）

版式基线（解析自定稿，复刻时不得更改）：
  A4 11906×16838；页边距 top/bottom=1440、left/right=1457、header/footer=720 twips
  P00 标题「工程管理部通知单」黑体14pt加粗居中
  P01 编号行 黑体12pt：`编号：  XAHYYX-26xxx` + 3×Tab + `填写日期：     YYYY 年 M 月 D 日`
  表0 6行×4列：全边框single、行高atLeast 654/636/636/636/5152/3229
      栏目：工程名称 / 摘由 / 主送·抄送 / 回复要求 / 通知内容 / 处理结果验证
      宋体12pt，栏目标签列加粗居中
  P02 填报说明行（表格**下方**）黑体12pt，「填报说明：」加粗
  P04 附件行 黑体12pt
  表1 照片：2行×3列无边框，列宽5.2cm、图片宽4.9cm、图注宋体10.5pt居中（紧凑省纸）
"""
import os
import sys
import shutil
from docx import Document
from docx.shared import Pt, Cm, Twips
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

SKILL_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEMPLATE = os.path.join(SKILL_DIR, 'assets', '工程管理部通知单_固定模板_2026-09-07.docx')

# ---------- 本单内容（改单子只改这里） ----------
NO_LINE = '编号：  XAHYYX-26038          \t\t\t填写日期：     2026 年 9 月 7 日'
PROJECT = '梧桐苑项目'
ABSTRACT = '1#楼西侧肥槽回填质量整改'
MAIN_TO = '陕西建工集团股份有限公司'
CC_TO = '西安恒泰建设项目管理有限公司'
REPLY = '☑请在 3 日内处理完毕，并书面回复。  □仅作通知，不需回复。'
NOTICE = [
    ('通知内容：', True),
    ('2026年9月7日项目工程管理部对1#楼西侧基槽回填工程进行质量检查，发现以下问题：', False),
    ('1#楼西侧肥槽回填一次性回填至设计标高，未分层回填、未分层压实；回填作业前未向监理单位报验，'
     '未履行报验程序，无法保证回填施工质量，存在回填土后期下沉、室外管线及地面开裂等质量风险。', False),
    ('整改要求：', False),
    ('1.立即停工整改。1#楼西侧肥槽回填作业立即停止，对已回填部位进行返工处理。', False),
    ('2.先试验、后回填。回填前先行组织地库外墙防水质量检查及淋水试验，确认外墙防水层施工质量合格、'
     '地库无渗漏后，方可进行回填作业，防止回填后出现渗漏、返工处理。', False),
    ('3.分层回填、分层压实。严格按规范要求分层回填、分层压实，控制每层虚铺厚度，逐层进行压实度检测，'
     '压实度合格后方可进行上层回填，规避后期下沉风险。', False),
    ('4.严格履行报验程序。回填施工前须向监理单位报验，经验收合格后方可施工；'
     '整改完成后于3日内书面回复本项目工程管理部，并附整改后照片。', False),
]
PHOTOS = [
    ('微信图片_20260907160340_1_532.jpg', '1#楼西侧肥槽一次性回填现状'),
    ('微信图片_20260907160344_2_532.jpg', '回填至设计标高，表面未压实'),
    ('微信图片_20260907162044_9_149.jpg', '回填土倒入肥槽堆积未碾压'),
]
# ---------------------------------------------

COL_W = [1396, 3589, 1522, 2852]
ROW_H = [654, 636, 636, 636, 5152, 3229]


def set_font(run, name='宋体', size=12, bold=False):
    run.font.name = name
    run._element.rPr.rFonts.set(qn('w:eastAsia'), name)
    run.font.size = Pt(size)
    run.font.bold = bold


def clear_paragraph(p, keep_ppr=True):
    """清空段落内容（保留段落属性）"""
    for child in list(p._p):
        if keep_ppr and child.tag == qn('w:pPr'):
            continue
        p._p.remove(child)


def write_paragraph(p, text, name='宋体', size=12, bold=False, align=None):
    clear_paragraph(p)
    if align is not None:
        p.alignment = align
    r = p.add_run(text)
    set_font(r, name, size, bold)
    return r


def fill_cell(cell, lines, vcenter=True):
    """lines = [(text, bold)]；复用已有段落，多余的删除"""
    if vcenter:
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    ps = cell.paragraphs
    for i, (txt, bold) in enumerate(lines):
        if i < len(ps):
            p = ps[i]
        else:
            p = cell.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing = 1.3
        write_paragraph(p, txt, '宋体', 12, bold)
    for j in range(len(lines), len(ps)):
        ps[j]._p.getparent().remove(ps[j]._p)


def build_from_template(out_path, photo_dir):
    """复制固定模板 → 替换文本与照片（格式100%保真）"""
    if not os.path.exists(TEMPLATE):
        raise SystemExit('固定模板缺失：%s' % TEMPLATE)
    shutil.copyfile(TEMPLATE, out_path)
    doc = Document(out_path)

    # P01 编号行
    write_paragraph(doc.paragraphs[1], NO_LINE, '黑体', 12)

    t = doc.tables[0]
    fill_cell(t.cell(0, 1), [(PROJECT, False)])
    fill_cell(t.cell(1, 1), [(ABSTRACT, False)])
    fill_cell(t.cell(2, 1), [(MAIN_TO, False)])
    fill_cell(t.cell(2, 3), [(CC_TO, False)])
    fill_cell(t.cell(3, 1), [(REPLY, False)])
    fill_cell(t.cell(4, 0), NOTICE, vcenter=False)

    # 附件照片（模板为 2行×3列；照片数须为3）
    if len(PHOTOS) != 3:
        raise SystemExit('模板照片区为3张，当前 %d 张，请用 --rebuild 重绘' % len(PHOTOS))
    pt = doc.tables[1]
    for i, (fn, cap) in enumerate(PHOTOS):
        pc = pt.cell(0, i).paragraphs[0]
        clear_paragraph(pc)
        pc.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pc.add_run().add_picture(os.path.join(photo_dir, fn), width=Cm(4.9))
        cc = pt.cell(1, i).paragraphs[0]
        cc.alignment = WD_ALIGN_PARAGRAPH.CENTER
        write_paragraph(cc, '图%d：%s' % (i + 1, cap), '宋体', 10.5)
    doc.save(out_path)
    return out_path


def para(doc, text, size=12, name='宋体', bold=False, align=None,
         space_after=6, line=1.4):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_after = Pt(space_after)
    pf.line_spacing = line
    if align is not None:
        pf.alignment = align
    r = p.add_run(text)
    set_font(r, name, size, bold)
    return p


def set_row_height(row, twips, rule='atLeast'):
    trPr = row._tr.get_or_add_trPr()
    th = OxmlElement('w:trHeight')
    th.set(qn('w:val'), str(twips))
    th.set(qn('w:hRule'), rule)
    trPr.append(th)


def cell_text(cell, text, size=12, name='宋体', bold=False,
              align=None, vcenter=True, line=1.3):
    if vcenter:
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    for i, ln in enumerate(text.split('\n')):
        p = cell.paragraphs[0] if i == 0 else cell.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing = line
        if align is not None:
            p.alignment = align
        if ln.startswith('【B】'):
            ln = ln[3:]
            b2 = True
        else:
            b2 = bold
        r = p.add_run(ln)
        set_font(r, name, size, b2)


def build_from_scratch(out_path, photo_dir):
    doc = Document()
    sec = doc.sections[0]
    sec.page_width = Twips(11906)
    sec.page_height = Twips(16838)
    sec.top_margin = Twips(1440)
    sec.bottom_margin = Twips(1440)
    sec.left_margin = Twips(1457)
    sec.right_margin = Twips(1457)

    para(doc, '工程管理部通知单', size=14, name='黑体', bold=True,
         align=WD_ALIGN_PARAGRAPH.CENTER, space_after=10)
    para(doc, NO_LINE, size=12, name='黑体', space_after=8)

    tbl = doc.add_table(rows=6, cols=4)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    tblPr = tbl._tbl.tblPr
    borders = OxmlElement('w:tblBorders')
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        e = OxmlElement('w:' + edge)
        e.set(qn('w:val'), 'single')
        e.set(qn('w:sz'), '4')
        e.set(qn('w:color'), '000000')
        borders.append(e)
    tblPr.append(borders)
    for i, w in enumerate(COL_W):
        for row in tbl.rows:
            row.cells[i].width = Twips(w)

    set_row_height(tbl.rows[0], ROW_H[0])
    cell_text(tbl.cell(0, 0), '工程名称', bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    cell_text(tbl.cell(0, 1).merge(tbl.cell(0, 3)), PROJECT,
              align=WD_ALIGN_PARAGRAPH.CENTER)
    set_row_height(tbl.rows[1], ROW_H[1])
    cell_text(tbl.cell(1, 0), '摘由', bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    cell_text(tbl.cell(1, 1).merge(tbl.cell(1, 3)), ABSTRACT)
    set_row_height(tbl.rows[2], ROW_H[2])
    cell_text(tbl.cell(2, 0), '主送', bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    cell_text(tbl.cell(2, 1), MAIN_TO, align=WD_ALIGN_PARAGRAPH.CENTER)
    cell_text(tbl.cell(2, 2), '抄送', bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    cell_text(tbl.cell(2, 3), CC_TO, align=WD_ALIGN_PARAGRAPH.CENTER)
    set_row_height(tbl.rows[3], ROW_H[3])
    cell_text(tbl.cell(3, 0), '回复要求', bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    cell_text(tbl.cell(3, 1).merge(tbl.cell(3, 3)), REPLY)
    set_row_height(tbl.rows[4], ROW_H[4])
    notice_txt = '\n'.join(('【B】' + t if b else t) for t, b in NOTICE)
    cell_text(tbl.cell(4, 0).merge(tbl.cell(4, 3)), notice_txt, vcenter=False)
    set_row_height(tbl.rows[5], ROW_H[5])
    cell_text(tbl.cell(5, 0).merge(tbl.cell(5, 3)),
              '处理结果验证：\n\n\n\n\n\n\n验证人：         '
              '日期：     年   月   日', vcenter=False)

    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    r = p.add_run('填报说明：')
    set_font(r, '黑体', 12, bold=True)
    r = p.add_run('本表一式三份，项目工程管理部、监理单位、施工单位各执一份。')
    set_font(r, '黑体', 12)

    p = doc.add_paragraph()
    p.add_run().add_break(WD_BREAK.PAGE)
    para(doc, '附件：现场质量检查照片（共 3 张，1#楼西侧肥槽回填）',
         size=12, name='黑体', space_after=8)
    pt = doc.add_table(rows=2, cols=3)
    pt.alignment = WD_TABLE_ALIGNMENT.CENTER
    pt.autofit = False
    for ci, (fn, cap) in enumerate(PHOTOS):
        for ri in range(2):
            pt.cell(ri, ci).width = Cm(5.2)
            pt.cell(ri, ci).vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        pc = pt.cell(0, ci).paragraphs[0]
        pc.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pc.add_run().add_picture(os.path.join(photo_dir, fn), width=Cm(4.9))
        cc = pt.cell(1, ci).paragraphs[0]
        cc.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = cc.add_run('图%d：%s' % (ci + 1, cap))
        set_font(r, '宋体', 10.5)
    doc.save(out_path)
    return out_path


def main():
    args = sys.argv[1:]
    rebuild = '--rebuild' in args
    photo_dir = os.getcwd()
    out = os.path.join(os.getcwd(), '工程管理部通知单.docx')
    if '--photo-dir' in args:
        photo_dir = args[args.index('--photo-dir') + 1]
    if '--out' in args:
        out = args[args.index('--out') + 1]
    fn = build_from_scratch if rebuild else build_from_template
    print('saved:', fn(out, photo_dir))


if __name__ == '__main__':
    main()
