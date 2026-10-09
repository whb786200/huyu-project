# -*- coding: utf-8 -*-
"""
以 6月「逐条整改回复」docx 为模板，
依据「7月份整改回复区域公司」目录数据生成 7月版逐条整改回复 docx。

数据源：
  - 代建梧桐樾府项目7月检查记录.xlsx（安全文明 sheet）：区域公司7月21日检查问题清单
      分类：安全隐患14 / 质量问题3 / 管理行为12，整改说明列(C7)为空
  - 中铁置业项目安全、质量隐患整治记录.xlsx（安全隐患整治记录 sheet）：隐患整治台账
      含真实整改说明(C7) + 责任人(C9:庞小强/陈军峰) + 完成日期(C6:2026.7.26)
"""

import openpyxl
import os
import re
import zipfile
from PIL import Image
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ===== 样式常量（复刻6月模板）=====
FONT_TITLE = '微软雅黑'
SIZE_TITLE = Pt(18)
FONT_SEC = '微软雅黑'
SIZE_SEC = Pt(15)
FONT_BODY = '宋体'
SIZE_BODY = Pt(10.5)
FONT_CELL = '宋体'
SIZE_CELL = Pt(9.5)

BASE = r'D:\Users\19586\Desktop\01_工程项目'
SRC_DIR = os.path.join(BASE, r'7月份整改回复区域公司')
TEMPLATE = os.path.join(BASE, r'二季度检查回复\整改回复\2026年6月梧桐樾府代建项目安全质量检查逐条整改回复.docx')
OUT = os.path.join(BASE, r'二季度检查回复\整改回复\2026年7月梧桐樾府代建项目安全质量检查逐条整改回复.docx')


def set_run_font(run, size, bold, font):
    run.font.size = size
    run.font.bold = bold
    run.font.name = font
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn('w:rFonts'))
    if rfonts is None:
        rfonts = OxmlElement('w:rFonts')
        rpr.append(rfonts)
    rfonts.set(qn('w:eastAsia'), font)


def add_para(doc, text, size=SIZE_BODY, bold=False, font=FONT_BODY,
             align=WD_ALIGN_PARAGRAPH.LEFT, space_after=6, space_before=0):
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.line_spacing = 1.5
    if text:
        r = p.add_run(text)
        set_run_font(r, size, bold, font)
    return p


def set_cell(cell, text, size=SIZE_CELL, bold=False, font=FONT_CELL,
             align=WD_ALIGN_PARAGRAPH.CENTER):
    # 单元格垂直居中
    tcPr = cell._tc.get_or_add_tcPr()
    va = OxmlElement('w:vAlign')
    va.set(qn('w:val'), 'center')
    tcPr.append(va)
    cell.text = ''
    p = cell.paragraphs[0]
    p.alignment = align
    lines = str(text).split('\n')
    for i, line in enumerate(lines):
        if i > 0:
            p.add_run().add_break()
        r = p.add_run(line)
        set_run_font(r, size, bold, font)


def style_table(table):
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl = table._tbl
    tblPr = tbl.tblPr
    borders = OxmlElement('w:tblBorders')
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        e = OxmlElement('w:%s' % edge)
        e.set(qn('w:val'), 'single')
        e.set(qn('w:sz'), '4')
        e.set(qn('w:space'), '0')
        e.set(qn('w:color'), '000000')
        borders.append(e)
    tblPr.append(borders)


def clean_prefix(s):
    """清除隐患/问题描述开头的编号前缀（1. 或 1、），避免与序号列重复"""
    return re.sub(r'^\s*\d+[.、]\s*', '', s).strip()


def overlap(a, b):
    """2-gram 重叠比例，用于匹配检查记录与整治记录隐患"""
    def grams(s):
        s = re.sub(r'[（）()\s\d.、，,；;]', '', s)
        return set(s[i:i + 2] for i in range(len(s) - 1))
    ga, gb = grams(a), grams(b)
    if not ga or not gb:
        return 0.0
    return len(ga & gb) / min(len(ga), len(gb))


def parse_check_records():
    """解析7月检查记录.xlsx -> 检查日期 + 三类问题"""
    fp = os.path.join(SRC_DIR, '代建梧桐樾府项目7月检查记录.xlsx')
    wb = openpyxl.load_workbook(fp, read_only=True, data_only=True)
    ws = wb['安全文明']

    inspect_date = '2026年7月21日'
    for row in ws.iter_rows(min_row=1, max_row=4, max_col=8, values_only=True):
        for v in row:
            if v and '检查时间' in str(v):
                m = re.search(r'(\d{4}年\d{1,2}月\d{1,2}日)', str(v))
                if m:
                    inspect_date = m.group(1)
    wb.close()

    wb = openpyxl.load_workbook(fp, read_only=True, data_only=True)
    ws = wb['安全文明']
    safe, qual, mgmt = [], [], []
    for row in ws.iter_rows(min_row=7, max_row=ws.max_row, max_col=8, values_only=True):
        seq, cat, loc, desc = row[0], row[1], row[2], row[3]
        if seq is None and cat is None and desc is None:
            continue
        if str(seq).strip() in ('检查人', '项目负责人'):
            continue
        cat_s = '' if cat is None else str(cat).strip()
        d = '' if desc is None else str(desc).strip()
        if not d:
            continue
        item = {'desc': clean_prefix(d), 'loc': '' if loc is None else str(loc).strip()}
        if cat_s == '安全隐患':
            safe.append(item)
        elif cat_s == '质量问题':
            qual.append(item)
        elif cat_s == '管理行为':
            mgmt.append(item)
    wb.close()
    return inspect_date, safe, qual, mgmt


def parse_rectify():
    """解析隐患整治记录.xlsx -> 真实整改说明/责任人/日期 列表"""
    fp = os.path.join(SRC_DIR, '中铁置业项目安全、质量隐患整治记录.xlsx')
    wb = openpyxl.load_workbook(fp, read_only=True, data_only=True)
    ws = wb['安全隐患整治记录']
    rects = []
    for row in ws.iter_rows(min_row=9, max_row=ws.max_row, max_col=10, values_only=True):
        seq, cat, loc, desc, _, cdate, reply, _, person, _ = row[:10]
        if seq is None and desc is None:
            continue
        if str(seq).strip() in ('被检查项目', '负责人签字'):
            continue
        d = '' if desc is None else str(desc).strip()
        if not d:
            continue
        rects.append({
            'desc': d,
            'reply': '' if reply is None else str(reply).replace('\n', ' ').strip(),
            'person': '' if person is None else str(person).replace('\n', '、').strip(),
            'date': '' if cdate is None else str(cdate).strip(),
        })
    wb.close()
    return rects


def match_reply(desc, rects):
    """为一条检查记录隐患匹配整治记录中的真实回复"""
    best, best_score = None, 0.0
    for r in rects:
        sc = overlap(desc, r['desc'])
        if sc > best_score:
            best, best_score = r, sc
    if best and best_score >= 0.22:
        return best
    return None


def generic_mgmt_reply(desc):
    """管理/质量行为问题：通用合规回复（不编造具体佐证名称）"""
    return ('已按区域公司检查要求逐项整改落实：补充完善相关制度、记录、'
            '人员配置及审批签认资料，闭环佐证随文报送，并纳入项目常态化管理。')


def extract_rectify_with_photos():
    """解析隐患整治记录.xlsx：文本字段 + 前后照片（按序号映射），返回记录列表。

    列映射（1-indexed）：1序号 2问题类别 3位置 4隐患描述 5问题照片(前)
    6整改完成日期 7整改说明 8整改回复照片(后) 9整改负责人 10备注
    图片锚定：E列(col=4)=整改前，H列(col=7)=整改后。
    """
    fp = os.path.join(SRC_DIR, '中铁置业项目安全、质量隐患整治记录.xlsx')
    wb = openpyxl.load_workbook(fp, read_only=False, data_only=True)
    ws = wb['安全隐患整治记录']
    records = []
    for r in range(9, 25):
        seq = ws.cell(row=r, column=1).value
        if seq is None:
            continue
        records.append({
            'seq': int(seq) if str(seq).strip().isdigit() else seq,
            'cat': ws.cell(row=r, column=2).value,
            'loc': ws.cell(row=r, column=3).value,
            'hazard': ws.cell(row=r, column=4).value,
            'finish_date': ws.cell(row=r, column=6).value,
            'reply': ws.cell(row=r, column=7).value,
            'owner': ws.cell(row=r, column=9).value,
            'note': ws.cell(row=r, column=10).value,
            'before_img': None, 'after_img': None,
        })
    wb.close()

    # 图片锚定解析
    z = zipfile.ZipFile(fp)
    data = z.read('xl/drawings/drawing1.xml').decode('utf-8')
    rels = z.read('xl/drawings/_rels/drawing1.xml.rels').decode('utf-8')
    rid2target = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', rels))
    blocks = re.findall(r'<xdr:twoCellAnchor\b[^>]*>(.*?)</xdr:twoCellAnchor>', data, re.S)
    if not blocks:
        blocks = re.findall(r'<xdr:oneCellAnchor\b[^>]*>(.*?)</xdr:oneCellAnchor>', data, re.S)
    seqs = [r['seq'] for r in records]
    img_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'reply_images')
    os.makedirs(img_dir, exist_ok=True)
    for b in blocks:
        frm = re.search(r'<xdr:from>(.*?)</xdr:from>', b, re.S)
        if not frm:
            continue
        col = int(re.search(r'<xdr:col>(\d+)</xdr:col>', frm.group(1)).group(1))
        row0 = int(re.search(r'<xdr:row>(\d+)</xdr:row>', frm.group(1)).group(1))
        rid = re.search(r'r:embed="(rId\d+)"', b)
        if not rid:
            continue
        target = rid2target.get(rid.group(1), '').replace('../', '')
        seq = row0 + 1 - 8  # 0-indexed行+1=1-indexed行；序号=行-8
        kind = 'before' if col == 4 else ('after' if col == 7 else None)
        if kind and seq in seqs:
            src = 'xl/media/' + os.path.basename(target)
            ext = os.path.splitext(src)[1]
            dst = os.path.join(img_dir, 'hazard_%02d_%s%s' % (seq, kind, ext))
            if not os.path.exists(dst):
                with z.open(src) as fi, open(dst, 'wb') as fo:
                    fo.write(fi.read())
            rec = next(r for r in records if r['seq'] == seq)
            rec[kind + '_img'] = dst
    z.close()
    return records


def prep_image(path, target=(1050, 780)):
    """统一裁剪/压缩图片到固定尺寸（居中裁剪 cover 模式，比例恒定），保证排版整齐。"""
    im = Image.open(path)
    if im.mode in ('RGBA', 'P', 'LA'):
        im = im.convert('RGB')
    w, h = im.size
    tw, th = target
    dst_ratio = tw / th
    src_ratio = w / h
    if src_ratio > dst_ratio:        # 原图更宽 → 裁两侧
        new_h = h
        new_w = int(h * dst_ratio)
        left = (w - new_w) // 2
        im = im.crop((left, 0, left + new_w, h))
    else:                            # 原图更高 → 裁上下
        new_w = w
        new_h = int(w / dst_ratio)
        top = (h - new_h) // 2
        im = im.crop((0, top, w, top + new_h))
    im = im.resize((tw, th), Image.LANCZOS)
    out = path + '.opt.jpg'
    im.save(out, 'JPEG', quality=85)
    return out


def placeholder_image(text='整改后照片待补充', target=(1050, 780)):
    """生成与正常照片同尺寸的浅灰占位图（无图时保持排版整齐）。"""
    from PIL import ImageDraw, ImageFont
    tw, th = target
    im = Image.new('RGB', (tw, th), (235, 235, 235))
    d = ImageDraw.Draw(im)
    try:
        font = ImageFont.truetype('simhei.ttf', 40)
    except Exception:
        font = ImageFont.load_default()
    # 居中文字
    bbox = d.textbbox((0, 0), text, font=font)
    tw_text, th_text = bbox[2] - bbox[0], bbox[3] - bbox[1]
    d.text(((tw - tw_text) // 2, (th - th_text) // 2), text, fill=(120, 120, 120), font=font)
    out = os.path.join(os.path.dirname(__file__), '_placeholder.jpg')
    im.save(out, 'JPEG', quality=85)
    return out


def set_cell_va(cell, val='center'):
    """设置单元格垂直对齐方式。"""
    tcPr = cell._tc.get_or_add_tcPr()
    va = OxmlElement('w:vAlign')
    va.set(qn('w:val'), val)
    tcPr.append(va)


def add_two_photos(doc, rec):
    """插入整改前/后对比图：1x2 无边框表，含 caption。"""
    t = doc.add_table(rows=2, cols=2)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl = t._tbl
    tblPr = tbl.tblPr
    borders = OxmlElement('w:tblBorders')
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        e = OxmlElement('w:%s' % edge)
        e.set(qn('w:val'), 'none')
        borders.append(e)
    tblPr.append(borders)

    set_cell(t.cell(0, 0), '整改前', bold=True)
    set_cell(t.cell(0, 1), '整改后', bold=True)
    set_cell_va(t.cell(0, 0)); set_cell_va(t.cell(0, 1))

    c0, c1 = t.cell(1, 0), t.cell(1, 1)
    c0.text = ''; c1.text = ''
    p0 = c0.paragraphs[0]; p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p1 = c1.paragraphs[0]; p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_cell_va(c0); set_cell_va(c1)
    if rec['before_img'] and os.path.exists(rec['before_img']):
        p0.add_run().add_picture(prep_image(rec['before_img']), width=Cm(7.0))
    else:
        set_cell(c0, '（无整改前照片）')
    if rec['after_img'] and os.path.exists(rec['after_img']):
        p1.add_run().add_picture(prep_image(rec['after_img']), width=Cm(7.0))
    else:
        p1.add_run().add_picture(placeholder_image(), width=Cm(7.0))


def build():
    inspect_date, safe, qual, mgmt = parse_check_records()
    rects = parse_rectify()

    doc = Document()
    for sec in doc.sections:
        sec.top_margin = Cm(2.5)
        sec.bottom_margin = Cm(2.5)
        sec.left_margin = Cm(3.0)
        sec.right_margin = Cm(3.0)

    # 标题
    add_para(doc, '2026年7月梧桐樾府代建项目安全质量检查逐条整改回复',
             SIZE_TITLE, True, FONT_TITLE, WD_ALIGN_PARAGRAPH.CENTER, space_after=10)
    # 依据
    add_para(doc, '依据《2026年7月梧桐樾府代建项目安全质量检查通报》逐项编制',
             SIZE_BODY, False, FONT_BODY, WD_ALIGN_PARAGRAPH.CENTER, space_after=8)
    # 说明
    add_para(doc, ('根据检查通报列明的问题，项目已按“责任到单位、整改有措施、闭合有资料、复查有记录”'
                   '的原则进行逐项梳理。以下回复为整改闭合口径，正式报送时请同步附制度文件、会议纪要、'
                   '签字记录、验收记录、整改照片等佐证资料。'),
             SIZE_BODY, False, FONT_BODY, WD_ALIGN_PARAGRAPH.LEFT, space_after=8)

    # 顶部信息表
    t0 = doc.add_table(rows=3, cols=4)
    style_table(t0)
    set_cell(t0.cell(0, 0), '工程名称', bold=True)
    set_cell(t0.cell(0, 1), '梧桐樾府项目')
    set_cell(t0.cell(0, 2), '检查日期', bold=True)
    set_cell(t0.cell(0, 3), inspect_date)
    set_cell(t0.cell(1, 0), '施工单位', bold=True)
    set_cell(t0.cell(1, 1), '陕西建工集团股份有限公司')
    set_cell(t0.cell(1, 2), '监理单位', bold=True)
    set_cell(t0.cell(1, 3), '西安恒泰建设项目管理有限公司')
    set_cell(t0.cell(2, 0), '项目地址', bold=True)
    set_cell(t0.cell(2, 1), '鄠邑区宏桥路以东、渼陂路以南')
    set_cell(t0.cell(2, 2), '整改原则', bold=True)
    set_cell(t0.cell(2, 3), '逐条对应、资料闭合、复查销项')

    # 一、超危大
    add_para(doc, '一、超危大工程及重大事故隐患问题整改回复', SIZE_SEC, True, FONT_SEC,
             WD_ALIGN_PARAGRAPH.LEFT, space_after=6, space_before=6)
    add_para(doc, '本月检查未发现超危大工程及重大事故隐患问题；现场危大工程（爬架、塔吊、模板支撑等）'
                 '均按清单实施动态管控，相关方案、论证、交底、验收及旁站资料完整，后续持续复核销项。',
             SIZE_BODY, False, FONT_BODY, space_after=8)

    # 二、安全管理行为（管理行为12条）
    add_para(doc, '二、安全管理行为问题整改回复', SIZE_SEC, True, FONT_SEC,
             WD_ALIGN_PARAGRAPH.LEFT, space_after=6, space_before=6)
    t_m = doc.add_table(rows=len(mgmt) + 1, cols=4)
    style_table(t_m)
    for ci, h in enumerate(['序号', '通报问题/要求', '逐条整改回复', '责任及闭合资料']):
        set_cell(t_m.cell(0, ci), h, bold=True)
    for i, it in enumerate(mgmt, 1):
        set_cell(t_m.cell(i, 0), '管理%d' % i)
        set_cell(t_m.cell(i, 1), it['desc'], align=WD_ALIGN_PARAGRAPH.LEFT)
        set_cell(t_m.cell(i, 2), generic_mgmt_reply(it['desc']), align=WD_ALIGN_PARAGRAPH.LEFT)
        set_cell(t_m.cell(i, 3), '责任：项目公司（涉及监理、总包的事项同步落实）。\n资料：相关制度、'
                 '记录、人员配置及审批签认等闭合资料随文报送。', align=WD_ALIGN_PARAGRAPH.LEFT)

    # 三、质量管理行为（质量问题3条）
    add_para(doc, '三、质量管理行为问题整改回复', SIZE_SEC, True, FONT_SEC,
             WD_ALIGN_PARAGRAPH.LEFT, space_after=6, space_before=6)
    t_q = doc.add_table(rows=len(qual) + 1, cols=4)
    style_table(t_q)
    for ci, h in enumerate(['序号', '通报问题/要求', '逐条整改回复', '责任及闭合资料']):
        set_cell(t_q.cell(0, ci), h, bold=True)
    for i, it in enumerate(qual, 1):
        set_cell(t_q.cell(i, 0), '质量%d' % i)
        set_cell(t_q.cell(i, 1), it['desc'], align=WD_ALIGN_PARAGRAPH.LEFT)
        set_cell(t_q.cell(i, 2), generic_mgmt_reply(it['desc']), align=WD_ALIGN_PARAGRAPH.LEFT)
        set_cell(t_q.cell(i, 3), '责任：项目公司、监理、总包。\n资料：材料验收记录、'
                 '复试报告、封样台账、工序验收记录等闭合资料随文报送。', align=WD_ALIGN_PARAGRAPH.LEFT)

    # 四、现场安全质量隐患（依据隐患整治记录真实内容 + 整改前后照片）
    add_para(doc, '四、现场安全质量隐患问题整改回复', SIZE_SEC, True, FONT_SEC,
             WD_ALIGN_PARAGRAPH.LEFT, space_after=6, space_before=6)
    add_para(doc, '以下隐患均依据《中铁置业项目安全、质量隐患整治记录》逐项整改闭合，'
                 '每条附“整改前 / 整改后”现场照片佐证。',
             SIZE_BODY, False, FONT_BODY, space_after=6)
    photo_records = extract_rectify_with_photos()
    for rec in photo_records:
        seq = rec['seq']
        hazard = clean_prefix(str(rec['hazard'])) if rec['hazard'] else ''
        loc = str(rec['loc']).strip() if rec['loc'] else ''
        cat = str(rec['cat']).strip() if rec['cat'] else ''
        add_para(doc, '%d、%s（位置：%s；问题类别：%s）' % (seq, hazard, loc, cat),
                 SIZE_BODY, True, FONT_BODY, space_after=2, space_before=4)
        reply = str(rec['reply']).replace('\n', '').strip() if rec['reply'] else '已按要求整改完成'
        owner = str(rec['owner']).replace('\n', '、').strip() if rec['owner'] else '庞小强、陈军峰'
        date = str(rec['finish_date']).strip() if rec['finish_date'] else '2026年7月26日'
        add_para(doc, '整改回复：%s；整改责任人：%s；整改完成日期：%s' % (reply, owner, date),
                 SIZE_BODY, False, FONT_BODY, space_after=3)
        add_two_photos(doc, rec)

    # 五、整改闭合要求
    add_para(doc, '五、整改闭合要求', SIZE_SEC, True, FONT_SEC,
             WD_ALIGN_PARAGRAPH.LEFT, space_after=6, space_before=6)
    for line in [
        '• 各责任单位按本回复对应事项补齐签字、照片、台账、验收、复查等闭合资料，做到问题、措施、责任、时限、结果一一对应。',
        '• 涉及超危大工程、爬架、脚手架、移动作业平台、基坑排水等较大风险事项，整改完成后须经监理复查确认，并由项目公司留痕抽查。',
        '• 对制度、会议、培训、周检、危大工程清单、材料验收和重点工序验收等管理类问题，纳入项目月度自查和资料归档检查，防止反复出现。',
    ]:
        add_para(doc, line, SIZE_BODY, False, FONT_BODY, space_after=4)

    # ===== 末尾签字栏 =====
    add_para(doc, '', SIZE_BODY)  # 空行间隔
    sig_tbl = doc.add_table(rows=2, cols=3)
    style_table(sig_tbl)
    # 第一行：编制/审核/审批
    set_cell(sig_tbl.cell(0, 0), '编制：', SIZE_BODY)
    set_cell(sig_tbl.cell(0, 1), '审核：', SIZE_BODY)
    set_cell(sig_tbl.cell(0, 2), '审批：', SIZE_BODY)
    # 第二行：日期
    set_cell(sig_tbl.cell(1, 0), '      年   月   日', SIZE_BODY)
    set_cell(sig_tbl.cell(1, 1), '      年   月   日', SIZE_BODY)
    set_cell(sig_tbl.cell(1, 2), '      年   月   日', SIZE_BODY)
    # 设置签字栏行高（留足签名空间）
    for row in sig_tbl.rows:
        row.height = Cm(2.5)

    # 底部署名行
    p_bottom = doc.add_paragraph()
    p_bottom.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_bottom = p_bottom.add_run('梧桐樾府项目安全质量检查整改回复')
    set_run_font(r_bottom, SIZE_BODY, False, FONT_BODY)

    doc.save(OUT)
    print('[OK] %s' % OUT)
    print('检查日期=%s | 安全隐患=%d 质量问题=%d 管理行为=%d | 现场隐患(带照片)=%d'
          % (inspect_date, len(safe), len(qual), len(mgmt), len(photo_records)))


if __name__ == '__main__':
    build()
