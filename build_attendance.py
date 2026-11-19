# -*- coding: utf-8 -*-
import re, json
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

SRC = r"D:\AI\2609_截10月31_原始檔案.xlsx"
OUT = r"D:\项目\AI\2609_出席統計_截至2026-10-31_raw.xlsx"

def dur(s):
    m = re.search(r"(\d{1,2}):(\d{2})-(\d{1,2}):(\d{2})", str(s))
    h1, m1, h2, m2 = map(int, m.groups())
    return (h2 * 60 + m2 - (h1 * 60 + m1)) / 60.0

def key(s):
    m = re.search(r"(\d{4})/(\d{2})/(\d{2})\s+(\d{1,2}):(\d{2})", str(s))
    return tuple(int(x) for x in m.groups())

wb_src = openpyxl.load_workbook(SRC, data_only=True)
ws_src = wb_src["sheet1"]

sessions = []
for c in range(2, 111):
    v = ws_src.cell(row=7, column=c).value
    sessions.append((c, str(v), dur(v), key(v)))

PRESENT = ("√", "△", "※")
ABSENT = ("×", "☆")
LATE = ("△",)
EARLY = ("※")

students = []
for r in range(8, 49):
    sid = ws_src.cell(row=r, column=1).value
    present = absent = late_h = early_h = 0.0
    events = []
    excused = 0
    for c, label, d, k in sessions:
        sym = ws_src.cell(row=r, column=c).value
        if sym in PRESENT:
            present += d
        elif sym in ABSENT:
            absent += d
        elif sym == "/":
            excused += 1
        if sym in PRESENT or sym in ABSENT:
            pass
        if sym in LATE or sym in EARLY:
            events.append({"label": label, "dur": d, "key": k,
                           "type": "遲到" if sym in LATE else "早退", "sym": sym})
    events.sort(key=lambda e: e["key"])
    total_le = sum(e["dur"] for e in events)
    n = len(events)
    checkpoint = (n // 3) * 3
    base = sum(e["dur"] for e in events[:checkpoint]) if checkpoint else 0.0
    students.append({
        "id": int(sid), "present": present, "absent": absent,
        "total_le": total_le, "base": base, "n": n, "excused": excused,
        "events": events, "checkpoint": checkpoint,
    })

# ---------- styles ----------
FONT = "Microsoft JhengHei"
thin = Side(style="thin", color="000000")
border = Border(left=thin, right=thin, top=thin, bottom=thin)
hdr_fill = PatternFill("solid", fgColor="F2F2F2")
hdr_font = Font(name=FONT, bold=True, size=11)
body_font = Font(name=FONT, size=11)
title_font = Font(name=FONT, bold=True, size=14)
center = Alignment(horizontal="center", vertical="center", wrap_text=True)

HEADERS = ["序號", "出席時數\n(未算遲到早退)", "缺席時數", "總遲到早退\n時數",
           "滿3次遲到\n早退時數", "因遲到早退\n需扣之時數",
           "總出席時數\n(已扣遲到早退)", "總教學時數", "總出席率%"]

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "出席統計"

ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(HEADERS))
t = ws.cell(row=1, column=1, value="2609 出席統計（截至 2026-10-31）")
t.font = title_font
t.alignment = Alignment(horizontal="left", vertical="center")
ws.row_dimensions[1].height = 30

for i, h in enumerate(HEADERS, 1):
    c = ws.cell(row=2, column=i, value=h)
    c.font = hdr_font; c.alignment = center; c.border = border; c.fill = hdr_fill
ws.row_dimensions[2].height = 48

for idx, s in enumerate(students):
    r = 3 + idx
    ws.cell(row=r, column=1, value=s["id"])
    ws.cell(row=r, column=2, value=round(s["present"], 6))
    ws.cell(row=r, column=3, value=round(s["absent"], 6))
    ws.cell(row=r, column=4, value=round(s["total_le"], 6))
    ws.cell(row=r, column=5, value=round(s["base"], 6))
    ws.cell(row=r, column=6, value="=E%d/3" % r)
    ws.cell(row=r, column=7, value="=B%d-F%d" % (r, r))
    ws.cell(row=r, column=8, value="=B%d+C%d" % (r, r))
    ws.cell(row=r, column=9, value='=IF(H%d=0,0,G%d/H%d)' % (r, r, r))
    for col in range(1, 10):
        cell = ws.cell(row=r, column=col)
        cell.border = border; cell.font = body_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        if col in (2, 3, 4, 5, 6, 7, 8):
            cell.number_format = "0.00"
        elif col == 9:
            cell.number_format = "0.0%"
        else:
            cell.number_format = "0"

widths = [7, 15, 11, 13, 13, 15, 17, 12, 11]
for i, w in enumerate(widths, 1):
    ws.column_dimensions[get_column_letter(i)].width = w
ws.freeze_panes = "B3"
ws.sheet_view.showGridLines = False

# ---------- 明細 ----------
ws2 = wb.create_sheet("遲到早退明細")
H2 = ["序號", "第幾次", "日期", "時間", "時數", "類型", "組別", "是否納入滿3次", "說明"]
for i, h in enumerate(H2, 1):
    c = ws2.cell(row=1, column=i, value=h)
    c.font = hdr_font; c.alignment = center; c.border = border; c.fill = hdr_fill
ws2.row_dimensions[1].height = 32
rr = 2
for s in students:
    for i, e in enumerate(s["events"], 1):
        date_s, time_s = e["label"].split("  ")
        included = i <= s["checkpoint"]
        group = (i + 2) // 3
        if included:
            note = "第%d組已滿3次，計入後按1/3扣除" % group
        else:
            note = "第%d組未滿3次，本次不罰" % group
        vals = [s["id"], i, date_s, time_s, round(e["dur"], 6), e["type"],
                "第%d組" % group, "是" if included else "否", note]
        for j, v in enumerate(vals, 1):
            c = ws2.cell(row=rr, column=j, value=v)
            c.font = body_font; c.border = border
            c.alignment = Alignment(horizontal="left" if j == 9 else "center", vertical="center")
            if j == 5:
                c.number_format = "0.00"
        rr += 1
for i, w in enumerate([7, 9, 12, 14, 9, 8, 9, 13, 34], 1):
    ws2.column_dimensions[get_column_letter(i)].width = w
ws2.freeze_panes = "A2"
ws2.auto_filter.ref = "A1:I%d" % (rr - 1)

# ---------- 計算說明 ----------
ws3 = wb.create_sheet("計算說明")
ws3.column_dimensions["A"].width = 26
ws3.column_dimensions["B"].width = 96
rows = [
    ("项目", "说明"),
    ("数据范围", "原始档 2609，课节 2026-09-02 至 2026-10-30，共 109 节；学生 41 人（序号 1–41）。"),
    ("签到符号", "√ 出席(准时)、△ 迟到、※ 早退、☆ 缺席(请假)、× 旷课、○ 豁免、/ 该节不排入此学生。"),
    ("出席时数(未算迟到早退)", "√、△、※ 三种状态的排课时长相加。迟到、早退先按整节计为出席，罚扣在后面的「需扣之时数」单独减。"),
    ("缺席时数", "× 旷课与 ☆ 请假的排课时长相加。"),
    ("总教学时数", "出席时数(未算迟到早退) + 缺席时数。全勤学生为 213.50 小时；序号 1 有 21 节「/」不排入，其总教学时数为 171.67。"),
    ("一次迟到或早退的时数", "原始表只有状态、没有迟到分钟，故每次按该节的「下课 − 上课」时长计算。"),
    ("排序", "同一学生的迟到、早退合并排序，按上课日期与时间由早到晚，依次编为第 1 次、第 2 次……"),
    ("满 3 次的处罚规则", "只处罚已满 3 的整数倍次数：第 1–3 次计入并按合计的 1/3 扣除（即只算 2/3）；第 4、5 次不罚；到第 6 次改为按第 1–6 次合计的 1/3 扣除；第 7、8 次不罚；第 9 次起按第 1–9 次合计，其后每满 3 次同样处理。"),
    ("满3次迟到早退时数", "已凑满的那些次的时数合计，即前 floor(次数/3)×3 次的时数相加；不足 3 次时为 0.00。"),
    ("因迟到早退需扣之时数", "满3次迟到早退时数 ÷ 3（保留原值，显示 2 位小数）。"),
    ("总出席时数(已扣迟到早退)", "出席时数(未算迟到早退) − 因迟到早退需扣之时数。"),
    ("总出席率%", "总出席时数(已扣) ÷ 总教学时数，以百分比显示 1 位小数。"),
    ("表格公式", "F = E/3；G = B − F；H = B + C；I = IF(H=0, 0, G/H)。B、C、D、E 为按原始签到统计的数值。"),
]
for i, (a, b) in enumerate(rows, 1):
    ca = ws3.cell(row=i, column=1, value=a)
    cb = ws3.cell(row=i, column=2, value=b)
    ca.font = Font(name=FONT, bold=(i == 1), size=11)
    cb.font = Font(name=FONT, bold=(i == 1), size=11)
    ca.alignment = Alignment(vertical="top", wrap_text=True)
    cb.alignment = Alignment(vertical="top", wrap_text=True)
    ca.border = border; cb.border = border
    if i == 1:
        ca.fill = hdr_fill; cb.fill = hdr_fill
ws3.sheet_view.showGridLines = False

wb.save(OUT)
print("SAVED", OUT)
print(json.dumps([{ "id": s["id"], "present": round(s["present"], 2), "absent": round(s["absent"], 2),
  "le": round(s["total_le"], 2), "base": round(s["base"], 2), "n": s["n"], "excused": s["excused"]} for s in students], ensure_ascii=False))

