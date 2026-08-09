import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.comments import Comment

wb = openpyxl.Workbook()

# ---- styles ----
FONT = "Arial"
def f(bold=False, color="000000", size=10):
    return Font(name=FONT, bold=bold, color=color, size=size)
BLUE = "0000FF"     # hardcoded inputs
BLACK = "000000"    # formulas
GREEN = "008000"    # cross-sheet links
title_font = Font(name=FONT, bold=True, size=14, color="FFFFFF")
hdr_font = Font(name=FONT, bold=True, size=10, color="FFFFFF")
hdr_fill = PatternFill("solid", fgColor="1F3864")
sub_fill = PatternFill("solid", fgColor="D9E1F2")
yellow = PatternFill("solid", fgColor="FFFF00")
tot_fill = PatternFill("solid", fgColor="E2EFDA")
warn_fill = PatternFill("solid", fgColor="FCE4D6")
ok_fill = PatternFill("solid", fgColor="C6EFCE")
thin = Side(style="thin", color="BFBFBF")
border = Border(left=thin, right=thin, top=thin, bottom=thin)
center = Alignment(horizontal="center", vertical="center")
left = Alignment(horizontal="left", vertical="center", wrap_text=True)
right = Alignment(horizontal="right", vertical="center")

CUR0 = '$#,##0;($#,##0);-'
CUR2 = '$#,##0.00;($#,##0.00);-'
PCT  = '0.0%'
PCT4 = '0.0000%'

def style_row_border(ws, row, c1, c2):
    for c in range(c1, c2+1):
        ws.cell(row=row, column=c).border = border

# ============================================================
# Sheet 1: 说明
# ============================================================
ws = wb.active
ws.title = "说明"
ws.sheet_view.showGridLines = False
ws.merge_cells("A1:D1")
ws["A1"] = "数据一致性核对模型  |  2026 年 7 月 1 日 – 7 月 31 日"
ws["A1"].font = title_font
ws["A1"].fill = hdr_fill
ws["A1"].alignment = center
ws.row_dimensions[1].height = 26

r = 3
ws.cell(r,1,"一、核算公式").font = f(bold=True, size=11)
r += 1
formulas = [
    ("① MTD Cost", "= Action Earnings + AF Cost   ⟹   Action Earnings = MTD Cost − AF Cost"),
    ("② 核减金额", "= AF Cost − (AF Cost + Action Earnings) × 优化目标%  = AF Cost − MTD Cost × 优化目标%"),
    ("③ AF Cost%", "= AF Cost / MTD Cost"),
]
for name, formula in formulas:
    ws.cell(r,1,name).font = f(bold=True)
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
    c = ws.cell(r,2,formula); c.font = f(); c.alignment = left
    r += 1

r += 1
ws.cell(r,1,"二、账号口径").font = f(bold=True, size=11)
r += 1
accts = [
    ("6704696 (尾号4696)", "CPS 预算 / GMV / 优化目标 10% —— 涉及核减（结果=月末核减）。广告主=ByteC Media LTD。"),
    ("6494046 (尾号4046)", "CPA 预算 / DNC / 优化目标 15% —— 不涉及核减（结果=-），但影响广告主预算分配。广告主=ByteC-RTA。"),
]
for name, desc in accts:
    ws.cell(r,1,name).font = f(bold=True)
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
    c = ws.cell(r,2,desc); c.font = f(); c.alignment = left
    r += 1

r += 1
ws.cell(r,1,"三、颜色图例（用于分析：改蓝色格→全表重算）").font = f(bold=True, size=11)
r += 1
legend = [
    ("蓝色数字", BLUE, "手工录入项 / 可调参数（AF Cost、MTD Cost、优化目标%、SubId2 收益）——分析时改这里"),
    ("黑色数字", BLACK, "公式计算结果（Action Earnings、AF Cost%、核减金额）——请勿手改"),
    ("绿色数字", GREEN, "跨工作表引用"),
    ("黄色底纹", None, "关键假设 / 需人工确认的单元格"),
]
for name, col, desc in legend:
    c = ws.cell(r,1,name)
    c.font = f(bold=True, color=col if col else BLACK)
    if col is None:
        c.fill = yellow
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
    d = ws.cell(r,2,desc); d.font = f(); d.alignment = left
    r += 1

r += 1
ws.cell(r,1,"四、工作表").font = f(bold=True, size=11)
r += 1
for name, desc in [
    ("4696-CPS核算", "ByteC Media LTD 各地区 AF Cost% / MTD / 核减金额（含公式）"),
    ("4046-CPA核算", "ByteC-RTA 各地区核算 + 与给定 $32,365 交叉验证"),
    ("SubId2明细-4696", "图2 平台原始数据，按 SubId2 归集并分 TH/PH/SEA 桶"),
    ("一致性核对", "三方交叉核对与 SEA 大盘分摊反推、差异定位"),
]:
    ws.cell(r,1,name).font = f(bold=True)
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
    d = ws.cell(r,2,desc); d.font = f(); d.alignment = left
    r += 1

ws.column_dimensions["A"].width = 22
for col in "BCDEFGH":
    ws.column_dimensions[col].width = 16

# ============================================================
# helper to build a 核算 sheet
# ============================================================
headers = ["Region","Partner Name","Partner ID","Partner Type","优化目标%",
           "AF Cost","MTD Cost","Action Earnings","AF Cost%","核减/超标金额","结果"]

def build_calc_sheet(title, rows, note_cell=None):
    ws = wb.create_sheet(title)
    ws.sheet_view.showGridLines = False
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
    ws.cell(1,1,title).font = title_font
    ws.cell(1,1).fill = hdr_fill
    ws.cell(1,1).alignment = center
    ws.row_dimensions[1].height = 24
    hr = 2
    for j, h in enumerate(headers, start=1):
        c = ws.cell(hr, j, h); c.font = hdr_font; c.fill = hdr_fill
        c.alignment = center; c.border = border
    first = hr + 1
    for i, row in enumerate(rows):
        rr = first + i
        region, pname, pid, ptype, target, afcost, mtd, result = row
        ws.cell(rr,1,region).font = f(); ws.cell(rr,1).alignment=center
        ws.cell(rr,2,pname).font = f()
        ws.cell(rr,3,pid).font = f(color=BLUE); ws.cell(rr,3).alignment=center
        ws.cell(rr,4,ptype).font = f(); ws.cell(rr,4).alignment=center
        # inputs (blue)
        c=ws.cell(rr,5,target); c.font=f(color=BLUE); c.number_format=PCT; c.alignment=center
        c=ws.cell(rr,6,afcost); c.font=f(color=BLUE); c.number_format=CUR0
        c=ws.cell(rr,7,mtd);    c.font=f(color=BLUE); c.number_format=CUR0
        # formulas (black)
        c=ws.cell(rr,8,f"=G{rr}-F{rr}");              c.font=f(); c.number_format=CUR0   # Action Earnings
        c=ws.cell(rr,9,f"=IF(G{rr}=0,0,F{rr}/G{rr})"); c.font=f(); c.number_format=PCT4  # AF Cost%
        c=ws.cell(rr,10,f"=F{rr}-G{rr}*E{rr}");        c.font=f(); c.number_format=CUR0   # 核减
        c=ws.cell(rr,11,result); c.font=f(); c.alignment=center
        style_row_border(ws, rr, 1, len(headers))
    last = first + len(rows) - 1
    # totals
    tr = last + 1
    ws.cell(tr,1,"合计").font=f(bold=True); ws.cell(tr,1).alignment=center
    for col_letter, col in [("F",6),("G",7),("H",8),("J",10)]:
        c=ws.cell(tr,col,f"=SUM({col_letter}{first}:{col_letter}{last})")
        c.font=f(bold=True); c.number_format=CUR0
    # AF Cost% total = total AF / total MTD
    c=ws.cell(tr,9,f"=IF(G{tr}=0,0,F{tr}/G{tr})"); c.font=f(bold=True); c.number_format=PCT4
    for col in range(1, len(headers)+1):
        ws.cell(tr,col).fill = tot_fill
        ws.cell(tr,col).border = border
    # widths
    widths = [9,16,12,12,10,11,12,15,12,15,11]
    for j,w in enumerate(widths, start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(j)].width = w
    return ws, first, last, tr

# ---- 4696 CPS ----
rows4696 = [
    ("TH","ByteC Media LTD","6704696","GMV",0.10,1065,9444,"月末核减"),
    ("PH","ByteC Media LTD","6704696","GMV",0.10,1183,5630,"月末核减"),
]
ws96, f96, l96, t96 = build_calc_sheet("4696-CPS核算", rows4696)
# note under table
nr = t96 + 2
ws96.cell(nr,1,"说明：").font=f(bold=True)
ws96.merge_cells(start_row=nr, start_column=2, end_row=nr, end_column=11)
ws96.cell(nr,2,"AF Cost 与 MTD Cost 为原汇总表录入项（蓝）；Action Earnings、AF Cost%、核减金额为公式（黑）。"
              "注意：本表 MTD 为原汇总表口径；与平台 SubId2 归集的地区 AE 存在差异，详见「一致性核对」③。").font=f(); ws96.cell(nr,2).alignment=left
ws96.cell(f96,6).comment = Comment("原表AF Cost%=0.112766696，对应AF Cost≈1065.03，此处取显示整数1065，故AF Cost%≈11.277%，尾差来自四舍五入。","核对")

# ---- 4046 CPA ----
rows4046 = [
    ("TH","ByteC-RTA","6494046","DNC",0.15,3925,36290,"-"),
    ("PH","ByteC-RTA","6494046","DNC",0.15,12,12,"-"),
    ("MY","ByteC-RTA","6494046","DNC",0.15,6,6,"-"),
    ("VN","ByteC-RTA","6494046","DNC",0.15,4,4,"-"),
]
ws46, f46, l46, t46 = build_calc_sheet("4046-CPA核算", rows4046)
# cross-check against given 32365
cr = t46 + 2
ws46.cell(cr,1,"交叉验证").font=f(bold=True, size=11)
cr += 1
ws46.cell(cr,1,"给定 7月4046 Action Earnings").font=f()
c=ws46.cell(cr,8,32365); c.font=f(color=BLUE); c.number_format=CUR0; c.fill=yellow
cr += 1
ws46.cell(cr,1,"模型合计 Action Earnings").font=f()
c=ws46.cell(cr,8,f"=H{t46}"); c.font=f(); c.number_format=CUR0
cr += 1
ws46.cell(cr,1,"差异").font=f(bold=True)
c=ws46.cell(cr,8,f"=H{cr-1}-H{cr-2}"); c.font=f(bold=True); c.number_format=CUR0
c2=ws46.cell(cr,9,f'=IF(ABS(H{cr})<0.5,"✅ 一致","⚠️ 不一致")'); c2.font=f(bold=True)
for rr in range(cr-3, cr+1):
    ws46.cell(rr,8).fill = tot_fill if rr!=cr-3 else yellow

# ============================================================
# Sheet: SubId2明细-4696  (image 2)
# ============================================================
ws2 = wb.create_sheet("SubId2明细-4696")
ws2.sheet_view.showGridLines = False
ws2.merge_cells("A1:F1")
ws2["A1"]="SubId2 明细（图2 · ByteC Media LTD · Tiktok Shop SEA · Jul 2026）"
ws2["A1"].font=title_font; ws2["A1"].fill=hdr_fill; ws2["A1"].alignment=center
ws2.row_dimensions[1].height=24
h2 = ["SubId2","Region归属","Raw Clicks","Actions","Sale Amount","Action Earnings"]
for j,h in enumerate(h2, start=1):
    c=ws2.cell(2,j,h); c.font=hdr_font; c.fill=hdr_fill; c.alignment=center; c.border=border
sub_rows = [
    ("MP_s_TTS-impact-SEA-cps-PDP_s_TH","TH",205329,109699,501349.81,5536.86),
    ("MP_s_TTS-impact-SEA-cps-PDP_s_PH","PH",389904,109848,395658.56,3102.46),
    ("NX_s_TTS_s_TH","TH",94894,39105,152164.39,1701.37),
    ("Xmob_s_TTS-impact-SEA-cps-PDP_s_PH","PH",311634,33320,129644.56,1172.17),
    ("Xmob_s_TTS-impact-SEA-cps-PDP_s_TH","TH",162372,14209,68108.96,759.17),
    ("KCAJ_s_TTS_s_TH","TH",96706,7078,33246.99,425.73),
    ("KCAJ_s_TTS_s_PH","PH",12,3860,9446.70,145.46),
    ("Bidmatrix_s_TTS_s_TH","TH",57845,4329,17993.87,95.51),
    ("Bidmatrix_s_TTS_s_PH","PH",32752,4685,13169.98,8.45),
]
fs = 3
for i,(sid,reg,rc,act,sale,ae) in enumerate(sub_rows):
    rr=fs+i
    ws2.cell(rr,1,sid).font=f()
    c=ws2.cell(rr,2,reg); c.font=f(color=BLUE); c.alignment=center
    c=ws2.cell(rr,3,rc); c.font=f(color=BLUE); c.number_format='#,##0'
    c=ws2.cell(rr,4,act); c.font=f(color=BLUE); c.number_format='#,##0'
    c=ws2.cell(rr,5,sale); c.font=f(color=BLUE); c.number_format=CUR2
    c=ws2.cell(rr,6,ae); c.font=f(color=BLUE); c.number_format=CUR2
    style_row_border(ws2, rr, 1, 6)
ls = fs+len(sub_rows)-1
tr2 = ls+1
ws2.cell(tr2,1,"平台合计").font=f(bold=True)
for col_letter,col,fmt in [("C",3,'#,##0'),("D",4,'#,##0'),("E",5,CUR2),("F",6,CUR2)]:
    c=ws2.cell(tr2,col,f"=SUM({col_letter}{fs}:{col_letter}{ls})"); c.font=f(bold=True); c.number_format=fmt
for col in range(1,7):
    ws2.cell(tr2,col).fill=tot_fill; ws2.cell(tr2,col).border=border

# buckets
br = tr2+2
ws2.cell(br,1,"按地区分桶（Action Earnings）").font=f(bold=True,size=11)
br+=1
ws2.cell(br,1,"（更新图：所有 SubId2 均带国家后缀 _TH/_PH，可直接按国家归集，无需再拆 SEA）").font=f(size=9,color="808080")
br+=1
buckets = [
    ("TH 合计", f'=SUMIF($B${fs}:$B${ls},"TH",$F${fs}:$F${ls})'),
    ("PH 合计", f'=SUMIF($B${fs}:$B${ls},"PH",$F${fs}:$F${ls})'),
]
bucket_rows = {}
for name,formula in buckets:
    ws2.cell(br,1,name).font=f(bold=True)
    c=ws2.cell(br,6,formula); c.font=f(bold=True); c.number_format=CUR2
    bucket_rows[name.split(" ")[0]] = br
    br+=1
ws2.cell(br,1,"合计（应=平台合计）").font=f(bold=True)
c=ws2.cell(br,6,f"=SUM(F{br-2}:F{br-1})"); c.font=f(bold=True); c.number_format=CUR2; c.fill=tot_fill
sum_check_row = br
ws2.column_dimensions["A"].width=38
ws2.column_dimensions["B"].width=12
for col in "CDEF":
    ws2.column_dimensions[col].width=15
TH_BUCKET=bucket_rows["TH"]; PH_BUCKET=bucket_rows["PH"]

# ============================================================
# Sheet: 一致性核对
# ============================================================
wc = wb.create_sheet("一致性核对")
wc.sheet_view.showGridLines=False
wc.merge_cells("A1:E1")
wc["A1"]="一致性核对与差异定位"
wc["A1"].font=title_font; wc["A1"].fill=hdr_fill; wc["A1"].alignment=center
wc.row_dimensions[1].height=24
R=3
def section(title):
    global R
    wc.cell(R,1,title).font=f(bold=True,size=11); wc.cell(R,1).fill=sub_fill
    for col in range(1,6): wc.cell(R,col).fill=sub_fill
    R+=1
def line(label, formula_or_val, fmt=CUR2, bold=False, link=False, note=""):
    global R
    wc.cell(R,1,label).font=f(bold=bold)
    c=wc.cell(R,3,formula_or_val)
    c.font=f(bold=bold, color=GREEN if link else BLACK)
    c.number_format=fmt
    if note:
        wc.cell(R,4,note).font=f(size=9,color="808080"); wc.cell(R,4).alignment=left
    R+=1
    return R-1

# --- 4046 ---
section("① 4046（CPA）—— 各地区 Action Earnings vs 给定 $32,365")
line("模型合计 Action Earnings", f"='4046-CPA核算'!H{t46}", CUR0, link=True)
line("给定值", 32365, CUR0)
d1=line("差异", f"=C{R-2}-C{R-1}", CUR0, bold=True)
wc.cell(d1,4,f'=IF(ABS(C{d1})<0.5,"✅ 完全一致","⚠️ 不一致")').font=f(bold=True)
R+=1

# --- 4696 platform total tie ---
section("② 4696（CPS）—— 平台合计 vs 汇总表 TH+PH 总额")
line("平台合计 Action Earnings（图2）", f"='SubId2明细-4696'!F{tr2}", CUR2, link=True)
line("汇总表 TH+PH（=MTD−AF）", f"='4696-CPS核算'!H{t96}", CUR2, link=True)
d2=line("总额差异", f"=C{R-2}-C{R-1}", CUR2, bold=True)
wc.cell(d2,4,"总额约差 $121，但需看到分地区才知问题").font=f(size=9,color="808080")
R+=1

# --- 4696 per-region comparison (THE core check) ---
section("③ 4696 分地区核对（更新图：SubId2 已带 _TH/_PH 后缀，可精确归集）★核心")
# header
hdr = ["地区","汇总表 AE(=MTD−AF)","平台 AE(SubId2)","AE 差异","校正 MTD(=平台AE+AF)","校正核减(AF−校正MTD×10%)","汇总表核减","核减差异"]
for j,htext in enumerate(hdr):
    c=wc.cell(R,1+j,htext); c.font=hdr_font; c.fill=hdr_fill; c.alignment=center; c.border=border
R+=1
region_rows=[("TH",f96,TH_BUCKET),("PH",f96+1,PH_BUCKET)]
cmp_first=R
for reg,srow,brow in region_rows:
    wc.cell(R,1,reg).font=f(bold=True); wc.cell(R,1).alignment=center
    c=wc.cell(R,2,f"='4696-CPS核算'!H{srow}"); c.font=f(color=GREEN); c.number_format=CUR2   # summary AE
    c=wc.cell(R,3,f"='SubId2明细-4696'!F{brow}"); c.font=f(color=GREEN); c.number_format=CUR2  # platform AE
    c=wc.cell(R,4,f"=B{R}-C{R}"); c.font=f(bold=True); c.number_format=CUR2                     # AE diff
    c=wc.cell(R,5,f"=C{R}+'4696-CPS核算'!F{srow}"); c.font=f(); c.number_format=CUR2            # corrected MTD
    c=wc.cell(R,6,f"='4696-CPS核算'!F{srow}-E{R}*'4696-CPS核算'!E{srow}"); c.font=f(); c.number_format=CUR2  # corrected 核减
    c=wc.cell(R,7,f"='4696-CPS核算'!J{srow}"); c.font=f(color=GREEN); c.number_format=CUR2      # summary 核减
    c=wc.cell(R,8,f"=G{R}-F{R}"); c.font=f(bold=True); c.number_format=CUR2                     # 核减 diff
    for col in range(1,9): wc.cell(R,col).border=border
    R+=1
# totals
wc.cell(R,1,"合计").font=f(bold=True); wc.cell(R,1).alignment=center
for col in [2,3,4,6,7,8]:
    cl=openpyxl.utils.get_column_letter(col)
    c=wc.cell(R,col,f"=SUM({cl}{cmp_first}:{cl}{R-1})"); c.font=f(bold=True); c.number_format=CUR2
for col in range(1,9): wc.cell(R,col).fill=tot_fill; wc.cell(R,col).border=border
cmp_tot=R
R+=2
wc.cell(R,1,"读法：").font=f(bold=True)
wc.merge_cells(start_row=R,start_column=2,end_row=R,end_column=8)
wc.cell(R,2,"「AE 差异」「核减差异」非 0 即为不一致。汇总表用的地区 AE 与平台 SubId2 归集不符，"
            "应以平台 AE 重算 MTD 与核减（见「校正」两列）。").font=f(); wc.cell(R,2).alignment=left
R+=2

# --- conclusion ---
section("④ 结论（基于更新图）")
concl = [
    ("4046（CPA）", "✅ 一致", "各地区AE合计=$32,365.00 精确吻合；公式全部对得上。"),
    ("4696 总额", "✅ 可对平", "平台合计$12,947.20；SubId2按_TH/_PH归集 TH=$8,518.64 + PH=$4,428.54 =$12,947.18。"),
    ("4696 分地区", "⚠️ 不一致", "汇总表 TH AE=8,379 vs 平台8,518.64（差−139.64）；PH AE=4,447 vs 平台4,428.54（差+18.46）。"),
    ("4696 核减金额", "⚠️ 需更正", "以平台AE重算：TH核减≈$107（原$121，多算约$14）、PH核减≈$622（原$620，少算约$2）。"),
]
for a,b,cc in concl:
    wc.cell(R,1,a).font=f(bold=True)
    c=wc.cell(R,2,b); c.font=f(bold=True)
    c.fill = ok_fill if "✅" in b else warn_fill
    wc.merge_cells(start_row=R, start_column=3, end_row=R, end_column=8)
    d=wc.cell(R,3,cc); d.font=f(); d.alignment=left
    R+=1
R+=1
wc.cell(R,1,"建议").font=f(bold=True)
R+=1
for s in [
    "1. 4696 的 TH/PH Action Earnings 应直接采用平台 SubId2 按 _TH/_PH 后缀归集的值（TH $8,518.64 / PH $4,428.54），据此重算 MTD 与核减金额。",
    "2. 复核汇总表 4696 的 MTD（TH 9,444 / PH 5,630）来源：与平台校正值（TH 9,583.64 / PH 5,611.54）不符，可能 MTD 中「商家&平台佣金」口径或取数时点与 Action Earnings 不一致。",
    "3. 表中若干 $1 级尾差（如4046 PH/VN的AF超标显示11/4）源于小额数据显示取整，非逻辑错误。",
]:
    wc.merge_cells(start_row=R, start_column=1, end_row=R, end_column=8)
    c=wc.cell(R,1,s); c.font=f(); c.alignment=left
    R+=1

widths=[8,20,18,13,22,26,14,13]
for j,w in enumerate(widths, start=1):
    wc.column_dimensions[openpyxl.utils.get_column_letter(j)].width=w

wb.save("/home/user/ByteC-OpenClaw/数据一致性核对_2026-07.xlsx")
print("saved")
