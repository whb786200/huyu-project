# -*- coding: utf-8 -*-
# 鄠邑梧桐樾府代建项目 —— 全员安全生产责任制考核评分表生成器
# 自动随机打分；财务部、营销部按更高分数区间评定（用户要求）。
# 一次生成多个月份，输出到指定文件夹。
# 运行：python gen_safety_assessment.py
import os
import random
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ===================== 可改常量 =====================
YEAR = 2026
MONTHS = [8, 9]  # 生成 8月、9月（可增删，如 [8,9,10]）
OUT_DIR = r"C:\Users\19586\Desktop\梧桐苑项目\01_工程项目\安全考核"
PROJECT = "梧桐苑项目（鄠邑梧桐樾府代建项目）"

# 考核维度（名称, 满分）
DIMS = [
    ("安全生产责任制建立与落实", 15),
    ("安全教育培训", 15),
    ("危险源辨识与隐患排查治理", 20),
    ("应急管理与演练", 15),
    ("现场安全管控", 20),
    ("安全投入与保障", 15),
]
TOTAL_MAX = sum(m for _, m in DIMS)  # 100

# 各维度随机区间：normal=一般部门，high=财务/营销（更高区间）
BANDS = {
    "normal": [(12, 15), (12, 15), (16, 20), (12, 15), (16, 20), (11, 15)],
    "high":   [(14, 15), (14, 15), (19, 20), (14, 15), (19, 20), (14, 15)],
}

# 考核对象名册：单位类别 | 部门 | 考核岗位/姓名 | 区间类型
# 代建项目部（中铁置业西安公司鄠邑项目部，5人真实）
# 监理部（西安恒泰，3人真实）
# 总包项目部（陕西建工，20人真实）
# 中铁置业西安公司六职能部门（部门级，负责人姓名待填，不虚构）
ROSTER = [
    ("代建项目部", "代建项目部", "李金良（项目总经理/组长）", "normal"),
    ("代建项目部", "代建项目部", "吴明乐（工程管理部部长）", "normal"),
    ("代建项目部", "代建项目部", "吴华兵（主管工程师）", "normal"),
    ("代建项目部", "代建项目部", "樊兆松（主管工程师）", "normal"),
    ("代建项目部", "代建项目部", "王成彧（综合管理员）", "normal"),
    ("监理部", "监理部", "陈军锋（项目总监）", "normal"),
    ("监理部", "监理部", "贾高（监理工程师）", "normal"),
    ("监理部", "监理部", "王伟（监理人员）", "normal"),
    ("总包项目部", "总包项目部", "庞小强（项目经理）", "normal"),
    ("总包项目部", "总包项目部", "赵建飞（安装经理）", "normal"),
    ("总包项目部", "总包项目部", "翟刘侃（生产经理）", "normal"),
    ("总包项目部", "总包项目部", "张宽（技术负责人）", "normal"),
    ("总包项目部", "总包项目部", "任强（安全总监）", "normal"),
    ("总包项目部", "总包项目部", "李宝恒（材料科长）", "normal"),
    ("总包项目部", "总包项目部", "刘玉波（造价员）", "normal"),
    ("总包项目部", "总包项目部", "易顺（材料员）", "normal"),
    ("总包项目部", "总包项目部", "殷乐意（施工员）", "normal"),
    ("总包项目部", "总包项目部", "杨利红（施工员）", "normal"),
    ("总包项目部", "总包项目部", "杨小明（施工员）", "normal"),
    ("总包项目部", "总包项目部", "李林刚（安装施工员）", "normal"),
    ("总包项目部", "总包项目部", "付跟社（安装施工员）", "normal"),
    ("总包项目部", "总包项目部", "王峰（安装施工员）", "normal"),
    ("总包项目部", "总包项目部", "董服义（安全员）", "normal"),
    ("总包项目部", "总包项目部", "赵哲（安全员）", "normal"),
    ("总包项目部", "总包项目部", "李正军（安全机管员）", "normal"),
    ("总包项目部", "总包项目部", "刘润龙（质量员）", "normal"),
    ("总包项目部", "总包项目部", "张少辉（试验员/资料主管）", "normal"),
    ("总包项目部", "总包项目部", "杨斐（资料员）", "normal"),
    ("中铁置业西安公司职能部门", "报批报建部", "李其蒙（报批报建部）", "normal"),
    ("中铁置业西安公司职能部门", "设计部", "韩东东（设计部）", "normal"),
    ("中铁置业西安公司职能部门", "营销部", "李炳龙（营销部）", "high"),
    ("中铁置业西安公司职能部门", "财务部", "王倩（财务部）", "high"),
    ("中铁置业西安公司职能部门", "成本部", "晁锁妮（成本部）", "normal"),
    ("中铁置业西安公司职能部门", "综合部", "部门负责人（待填）", "normal"),
]

# ===================== 样式 =====================
C_TITLE = Font(name="黑体", size=16, bold=True, color="FFFFFF")
C_HEAD = Font(name="黑体", size=10.5, bold=True, color="FFFFFF")
C_BODY = Font(name="宋体", size=10.5)
C_BODY_B = Font(name="宋体", size=10.5, bold=True)
HEAD_FILL = PatternFill("solid", fgColor="4472C4")
HIGH_FILL = PatternFill("solid", fgColor="FFF2CC")   # 财务/营销 高亮
TITLE_FILL = PatternFill("solid", fgColor="1F4E78")
EXC_FILL = PatternFill("solid", fgColor="C6EFCE")    # 优秀
GOOD_FILL = PatternFill("solid", fgColor="FFEB9C")   # 良好
AL_C = Alignment(horizontal="center", vertical="center", wrap_text=True)
AL_L = Alignment(horizontal="left", vertical="center", wrap_text=True)
THIN = Side(style="thin", color="B0B0B0")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def grade(t):
    if t >= 90:
        return "优秀", EXC_FILL
    if t >= 80:
        return "良好", GOOD_FILL
    if t >= 70:
        return "合格", None
    return "不合格", None


def build_one(seed, assess_date, out_file):
    rnd = random.Random(seed)
    wb = Workbook()

    # ---------- Sheet1 评分明细 ----------
    ws = wb.active
    ws.title = "评分明细"
    cols = ["序号", "单位类别", "部门", "考核岗位/姓名"] + [d[0] for d in DIMS] + ["合计(100)", "考核等级", "备注"]
    ncol = len(cols)

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncol)
    t = ws.cell(row=1, column=1, value="%s 全员安全生产责任制考核评分表（%s）" % (PROJECT, assess_date))
    t.font = C_TITLE
    t.fill = TITLE_FILL
    t.alignment = AL_C
    ws.row_dimensions[1].height = 30

    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ncol)
    s = ws.cell(row=2, column=1, value="评分方式：各维度随机评定；财务部、营销部按更高分数区间考核（黄底标识）。满分 %d 分。" % TOTAL_MAX)
    s.font = Font(name="宋体", size=9, italic=True, color="555555")
    s.alignment = AL_L
    ws.row_dimensions[2].height = 18

    hr = 3
    for j, c in enumerate(cols, 1):
        cell = ws.cell(row=hr, column=j, value=c)
        cell.font = C_HEAD
        cell.fill = HEAD_FILL
        cell.alignment = AL_C
        cell.border = BORDER
    ws.row_dimensions[hr].height = 34

    r = hr + 1
    for i, (cat, dept, name, band) in enumerate(ROSTER, 1):
        scores = [rnd.randint(lo, hi) for lo, hi in BANDS[band]]
        total = sum(scores)
        g, gfill = grade(total)
        note = "财务/营销 高区间" if band == "high" else ""
        row_vals = [i, cat, dept, name] + scores + [total, g, note]
        for j, v in enumerate(row_vals, 1):
            cell = ws.cell(row=r, column=j, value=v)
            cell.border = BORDER
            if j in (1, 2, 3):
                cell.font = C_BODY
                cell.alignment = AL_C
            elif j == 4:
                cell.font = C_BODY
                cell.alignment = AL_L
            else:
                cell.font = C_BODY
                cell.alignment = AL_C
            if band == "high":
                cell.fill = HIGH_FILL
        gcell = ws.cell(row=r, column=ncol - 1)
        if gfill:
            gcell.fill = gfill
            gcell.font = C_BODY_B
        ws.row_dimensions[r].height = 20
        r += 1

    widths = [5, 22, 14, 26] + [11] * len(DIMS) + [10, 9, 16]
    for j, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.freeze_panes = "A4"

    # ---------- Sheet2 部门考核汇总 ----------
    ws2 = wb.create_sheet("部门考核汇总")
    agg = {}
    rnd2 = random.Random(seed)
    detail = []
    for cat, dept, name, band in ROSTER:
        scores = [rnd2.randint(lo, hi) for lo, hi in BANDS[band]]
        detail.append((cat, dept, name, band, sum(scores)))
    for cat, dept, name, band, tot in detail:
        key = (cat, dept)
        d = agg.setdefault(key, {"n": 0, "sum": 0, "mx": -1, "mn": 999, "high": band == "high"})
        d["n"] += 1
        d["sum"] += tot
        d["mx"] = max(d["mx"], tot)
        d["mn"] = min(d["mn"], tot)

    cols2 = ["单位类别", "部门", "考核人数", "平均分", "最高分", "最低分", "考核等级", "备注"]
    ws2.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(cols2))
    t2 = ws2.cell(row=1, column=1, value="部门/单位考核汇总（按平均分降序）")
    t2.font = C_TITLE
    t2.fill = TITLE_FILL
    t2.alignment = AL_C
    ws2.row_dimensions[1].height = 28
    for j, c in enumerate(cols2, 1):
        cell = ws2.cell(row=2, column=j, value=c)
        cell.font = C_HEAD
        cell.fill = HEAD_FILL
        cell.alignment = AL_C
        cell.border = BORDER
    ws2.row_dimensions[2].height = 26

    rows2 = []
    for (cat, dept), d in agg.items():
        avg = round(d["sum"] / d["n"], 1)
        rows2.append((cat, dept, d["n"], avg, d["mx"], d["mn"], d["high"]))
    rows2.sort(key=lambda x: -x[3])

    rr = 3
    for cat, dept, n, avg, mx, mn, high in rows2:
        g, gfill = grade(avg)
        vals = [cat, dept, n, avg, mx, mn, g, "财务/营销 高区间" if high else ""]
        for j, v in enumerate(vals, 1):
            cell = ws2.cell(row=rr, column=j, value=v)
            cell.border = BORDER
            cell.font = C_BODY if j != 7 else C_BODY_B
            cell.alignment = AL_C if j != 1 else AL_L
            if high:
                cell.fill = HIGH_FILL
        if gfill:
            ws2.cell(row=rr, column=7).fill = gfill
        ws2.row_dimensions[rr].height = 20
        rr += 1
    for j, w in enumerate([24, 16, 10, 10, 10, 10, 10, 18], 1):
        ws2.column_dimensions[get_column_letter(j)].width = w
    ws2.freeze_panes = "A3"

    # ---------- Sheet3 评分标准 ----------
    ws3 = wb.create_sheet("评分标准")
    ws3.merge_cells(start_row=1, start_column=1, end_row=1, end_column=3)
    t3 = ws3.cell(row=1, column=1, value="考核评分标准与说明")
    t3.font = C_TITLE
    t3.fill = TITLE_FILL
    t3.alignment = AL_C
    ws3.row_dimensions[1].height = 28

    std = [
        ("考核维度", "满分", "评分要点"),
        ("安全生产责任制建立与落实", 15, "是否逐级签订责任书、责任清单到岗到人、履职记录完整"),
        ("安全教育培训", 15, "三级教育、班前讲话、专题培训覆盖率与考试合格率"),
        ("危险源辨识与隐患排查治理", 20, "风险分级管控、隐患台账、整改闭环率、重大隐患销项"),
        ("应急管理与演练", 15, "预案完备性、演练频次与效果、应急物资配置"),
        ("现场安全管控", 20, "危大工程、临边防护、机械设备、消防与用电管理"),
        ("安全投入与保障", 15, "安措费用提取与使用、保险与防护装备保障"),
        ("合计", TOTAL_MAX, "六维度加总，满分 %d 分" % TOTAL_MAX),
    ]
    for i, row in enumerate(std, 2):
        for j, v in enumerate(row, 1):
            cell = ws3.cell(row=i, column=j, value=v)
            cell.border = BORDER
            cell.alignment = AL_L if j == 3 else AL_C
            if i == 2:
                cell.font = C_HEAD
                cell.fill = HEAD_FILL
            else:
                cell.font = C_BODY_B if row[0] in ("合计",) else C_BODY
    base = len(std) + 2
    ws3.cell(row=base, column=1, value="考核等级").font = C_HEAD
    ws3.cell(row=base, column=1).fill = HEAD_FILL
    ws3.cell(row=base, column=1).alignment = AL_C
    ws3.cell(row=base, column=2, value="分数区间").font = C_HEAD
    ws3.cell(row=base, column=2).fill = HEAD_FILL
    ws3.cell(row=base, column=2).alignment = AL_C
    ws3.cell(row=base, column=3, value="说明").font = C_HEAD
    ws3.cell(row=base, column=3).fill = HEAD_FILL
    ws3.cell(row=base, column=3).alignment = AL_C
    levels = [
        ("优秀", "≥90 分", "责任制落实到位，予以表扬"),
        ("良好", "80–89 分", "总体达标，持续改进"),
        ("合格", "70–79 分", "基本达标，需补强薄弱环节"),
        ("不合格", "<70 分", "限期整改并复查"),
    ]
    for k, (g_, rng, desc) in enumerate(levels, base + 1):
        ws3.cell(row=k, column=1, value=g_).font = C_BODY_B
        ws3.cell(row=k, column=1).alignment = AL_C
        ws3.cell(row=k, column=2, value=rng).alignment = AL_C
        ws3.cell(row=k, column=3, value=desc).alignment = AL_L
        for j in (1, 2, 3):
            ws3.cell(row=k, column=j).border = BORDER
    note_r = base + len(levels) + 2
    notes = [
        "1. 本表覆盖代建项目部、监理部、总包项目部及中铁置业西安公司六职能部门，共 %d 个考核对象，即“全员”。" % len(ROSTER),
        "2. 各维度分数由程序按设定区间随机生成（随机种子=%d，可复现；修改种子可换一批分数）。" % seed,
        "3. 按考核要求，财务部、营销部采用更高分数区间评定（明细表中黄底标识），其余部门按正常区间。",
        "4. 中铁置业西安公司职能部门考核人：李其蒙（报批报建）、韩东东（设计）、李炳龙（营销）、王倩（财务）、晁锁妮（成本），均为项目考核记录在册人员；综合部负责人待填，未虚构人员。",
        "5. 代建项目部、监理部、总包项目部人员均为项目真实在册人员。",
    ]
    for k, txt in enumerate(notes):
        c = ws3.cell(row=note_r + k, column=1, value=txt)
        c.font = Font(name="宋体", size=9, color="333333")
        c.alignment = AL_L
        ws3.merge_cells(start_row=note_r + k, start_column=1, end_row=note_r + k, end_column=3)
    for j, w in enumerate([26, 12, 70], 1):
        ws3.column_dimensions[get_column_letter(j)].width = w

    os.makedirs(OUT_DIR, exist_ok=True)
    wb.save(out_file)
    print("SAVED: " + out_file)
    return detail


if __name__ == "__main__":
    for m in MONTHS:
        seed = YEAR * 100 + m
        label = "%d年%d月" % (YEAR, m)
        out_file = os.path.join(OUT_DIR, "全员安全生产责任制考核评分表_%s.xlsx" % label)
        detail = build_one(seed, label, out_file)
        hi = [t for c, d, n, b, t in detail if b == "high"]
        no = [t for c, d, n, b, t in detail if b != "high"]
        import statistics as st
        print("  %s -> 财务/营销 avg=%.1f (min%d/max%d) | 其余 avg=%.1f (min%d/max%d)" % (
            label, st.mean(hi), min(hi), max(hi), st.mean(no), min(no), max(no)))
