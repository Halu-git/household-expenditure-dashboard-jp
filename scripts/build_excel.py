# -*- coding: utf-8 -*-
"""
Excelダッシュボード生成スクリプト
サンプル第2弾A「日本の家計支出ダッシュボード」

入力: data/processed/ の整形済みCSV 7点
出力: output/household_expenditure_dashboard.xlsx

方針(設計書 design.md §4・§7):
- シート1・3・4は用途分類、シート2は品目分類。分類を混ぜない。
- 各ダッシュボードシートに数値表+Excelネイティブグラフ+出典・注意を配置。
- データシートにはCSVをそのまま投入する(値の手直し禁止)。
"""
import csv
import os
from collections import OrderedDict

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PROC = os.path.join(BASE, "data", "processed")
OUT = os.path.join(BASE, "output", "household_expenditure_dashboard.xlsx")
SRC_DATE = "取得日: 2026-11-12(e-Stat 公開ダウンロード)"

TITLE_FONT = Font(name="Meiryo", size=14, bold=True)
SUB_FONT = Font(name="Meiryo", size=10, bold=True)
NOTE_FONT = Font(name="Meiryo", size=9, color="404040")
HDR_FILL = PatternFill("solid", fgColor="2F5597")
HDR_FONT = Font(name="Meiryo", size=10, bold=True, color="FFFFFF")
CELL_FONT = Font(name="Meiryo", size=10)
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
FMT_INT = '#,##0'
FMT_PCT1 = '0.0'
FMT_DEC2 = '0.00'

L1_USE = ["食料", "住居", "光熱・水道", "家具・家事用品", "被服及び履物",
          "保健医療", "交通・通信", "教育", "教養娯楽", "その他の消費支出"]
REGIONS = ["北海道", "東北", "関東", "北陸", "東海", "近畿", "中国", "四国", "九州", "沖縄"]
CITY_CLASSES = ["大都市", "中都市", "小都市A", "小都市B・町村"]
AGE_GROUPS = ["～34歳", "35～39歳", "40～44歳", "45～49歳", "50～54歳", "55～59歳",
              "60～64歳", "65～69歳", "70～74歳", "75～79歳", "80～84歳", "85歳～"]


def load(name):
    with open(os.path.join(PROC, name), newline="", encoding="utf-8") as fh:
        return list(csv.DictReader(fh))


def fnum(v):
    if v is None or v == "":
        return None
    try:
        return float(v)
    except ValueError:
        return None


def put_title(ws, title, subtitle, source_line):
    ws["A1"] = title
    ws["A1"].font = TITLE_FONT
    ws["A2"] = subtitle
    ws["A2"].font = SUB_FONT
    ws["A3"] = source_line
    ws["A3"].font = NOTE_FONT


def put_notes(ws, row, notes):
    for i, n in enumerate(notes):
        c = ws.cell(row=row + i, column=1, value=n)
        c.font = NOTE_FONT
    return row + len(notes)


def write_table(ws, start_row, start_col, headers, rows, fmts=None, widths=None):
    for j, h in enumerate(headers):
        c = ws.cell(row=start_row, column=start_col + j, value=h)
        c.fill = HDR_FILL
        c.font = HDR_FONT
        c.border = BORDER
        c.alignment = Alignment(horizontal="center", wrap_text=True)
    for i, r in enumerate(rows):
        for j, v in enumerate(r):
            c = ws.cell(row=start_row + 1 + i, column=start_col + j, value=v)
            c.font = CELL_FONT
            c.border = BORDER
            if fmts and j < len(fmts) and fmts[j]:
                c.number_format = fmts[j]
    if widths:
        for j, w in enumerate(widths):
            ws.column_dimensions[get_column_letter(start_col + j)].width = w
    return start_row + len(rows)


use_annual = load("use_annual.csv")
item_annual = load("item_annual.csv")
item_monthly = load("item_monthly.csv")
basic_sel = load("basic_selective_annual.csv")
region_use = load("region_use.csv")
age_use = load("age_use.csv")
house_attr = load("household_attributes.csv")


def use_val(year, item):
    for r in use_annual:
        if r["year"] == str(year) and r["item_name"] == item:
            return fnum(r["value"])
    return None


def region_val(year, region, item):
    for r in region_use:
        if r["year"] == str(year) and r["region"] == region and r["item_name"] == item:
            return fnum(r["amount_yen_per_month"])
    return None


def age_val(year, age, item):
    for r in age_use:
        if r["year"] == str(year) and r["age_group"] == age and r["item_name"] == item:
            return fnum(r["amount_yen_per_month"])
    return None


def attr_val(year, attr):
    for r in house_attr:
        if (r["year"] == str(year) and r["dimension"] == "地域"
                and r["group_name"] == "全国" and r["attribute"] == attr):
            return fnum(r["value"])
    return None


wb = Workbook()

# =====================================================================
# シート1: KPI概要(用途分類)
# =====================================================================
ws = wb.active
ws.title = "1_KPI概要"
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.page_setup.orientation = "landscape"
ws.page_setup.paperSize = ws.PAPERSIZE_A4
put_title(
    ws,
    "日本の家計支出ダッシュボード — 1. KPI概要",
    "対象: 二人以上の世帯 / 分類: 用途分類(全国)",
    "出典: 総務省統計局「家計調査(家計収支編)」(政府統計コード 00200561) 長期時系列表2・表4、詳細結果表 / " + SRC_DATE,
)

basic_share = {}
for r in basic_sel:
    basic_share[(r["year"], r["classification_series"])] = fnum(r["basic_share_pct_calc"])
kpi_rows = [
    ["消費支出(円/月)", use_val(2015, "消費支出"), use_val(2020, "消費支出"), use_val(2025, "消費支出"),
     use_val(2025, "消費支出") - use_val(2015, "消費支出")],
    ["エンゲル係数(％)", use_val(2015, "エンゲル係数(％)"), use_val(2020, "エンゲル係数(％)"),
     use_val(2025, "エンゲル係数(％)"), use_val(2025, "エンゲル係数(％)") - use_val(2015, "エンゲル係数(％)")],
    ["基礎的支出比率(％、算出値)※1", basic_share[("2015", "2015年改定基準")],
     basic_share[("2020", "2020年改定基準")], basic_share[("2025", "2025年改定基準")], None],
    ["世帯人員(人)", attr_val(2015, "世帯人員(人)"), attr_val(2020, "世帯人員(人)"), attr_val(2025, "世帯人員(人)"),
     attr_val(2025, "世帯人員(人)") - attr_val(2015, "世帯人員(人)")],
    ["世帯主の平均年齢(歳)", attr_val(2015, "世帯主の年齢(歳)"), attr_val(2020, "世帯主の年齢(歳)"),
     attr_val(2025, "世帯主の年齢(歳)"), attr_val(2025, "世帯主の年齢(歳)") - attr_val(2015, "世帯主の年齢(歳)")],
    ["持家率(％)", attr_val(2015, "持家率(％)"), attr_val(2020, "持家率(％)"), attr_val(2025, "持家率(％)"),
     attr_val(2025, "持家率(％)") - attr_val(2015, "持家率(％)")],
]
r_end = write_table(ws, 5, 1, ["KPI", "2015年", "2020年", "2025年", "変化(2015→2025)"],
                    kpi_rows, fmts=[None, FMT_INT, FMT_INT, FMT_INT, FMT_INT],
                    widths=[30, 14, 14, 14, 18])
for rr in (7, 8, 11):
    for cc in (2, 3, 4, 5):
        ws.cell(row=rr, column=cc).number_format = FMT_PCT1
for cc in (2, 3, 4, 5):
    ws.cell(row=9, column=cc).number_format = FMT_DEC2

tr_years = list(range(2000, 2026))
tr_rows = [[y, use_val(y, "消費支出")] for y in tr_years]
tr_start = r_end + 2
ws.cell(row=tr_start - 1, column=1, value="消費支出の推移(円/月、2000〜2025年)").font = SUB_FONT
tr_end = write_table(ws, tr_start, 1, ["年", "消費支出"], tr_rows, fmts=[FMT_INT, FMT_INT])
ch = LineChart()
ch.title = "消費支出の推移(二人以上の世帯、円/月)"
ch.y_axis.title = "円/月"
ch.x_axis.title = "年"
ch.add_data(Reference(ws, min_col=2, min_row=tr_start, max_row=tr_end), titles_from_data=True)
ch.set_categories(Reference(ws, min_col=1, min_row=tr_start + 1, max_row=tr_end))
ch.height, ch.width = 7.5, 15
ch.legend = None
ws.add_chart(ch, "D%d" % tr_start)

comp_rows = []
tot15, tot25 = use_val(2015, "消費支出"), use_val(2025, "消費支出")
for item in L1_USE:
    comp_rows.append([item,
                      round(use_val(2015, item) / tot15 * 100, 1),
                      round(use_val(2025, item) / tot25 * 100, 1)])
c_start = tr_end + 2
ws.cell(row=c_start - 1, column=1, value="大分類10費目の構成比(％、用途分類・算出値)").font = SUB_FONT
c_end = write_table(ws, c_start, 1, ["費目", "2015年", "2025年"], comp_rows,
                    fmts=[None, FMT_PCT1, FMT_PCT1])
ch2 = BarChart()
ch2.type = "bar"
ch2.title = "費目別構成比(％) 2025年"
ch2.add_data(Reference(ws, min_col=3, min_row=c_start, max_row=c_end), titles_from_data=True)
ch2.set_categories(Reference(ws, min_col=1, min_row=c_start + 1, max_row=c_end))
ch2.height, ch2.width = 7.5, 13
ch2.legend = None
ws.add_chart(ch2, "D%d" % c_start)
ch3 = BarChart()
ch3.title = "構成比の比較 2015年 vs 2025年(％)"
ch3.add_data(Reference(ws, min_col=2, min_row=c_start, max_row=c_end), titles_from_data=True)
ch3.set_categories(Reference(ws, min_col=1, min_row=c_start + 1, max_row=c_end))
ch3.height, ch3.width = 7.5, 15
ws.add_chart(ch3, "L%d" % c_start)

series_rank = {"2010年改定前": 0, "2010年改定基準": 1, "2015年改定基準": 2,
               "2020年改定基準": 3, "2025年改定基準": 4}
best = {}
for r in basic_sel:
    y = int(r["year"])
    if y not in best or series_rank[r["classification_series"]] > series_rank[best[y]["classification_series"]]:
        best[y] = r
bs_rows = [[y, fnum(best[y]["basic_yen_per_month"]), fnum(best[y]["selective_yen_per_month"]),
            best[y]["classification_series"]] for y in sorted(best)]
b_start = c_end + 2
ws.cell(row=b_start - 1, column=1,
        value="基礎的支出・選択的支出(円/月)。改定年は新基準の値を採用し、系列の切替を併記").font = SUB_FONT
b_end = write_table(ws, b_start, 1, ["年", "基礎的支出", "選択的支出", "採用した基準系列"], bs_rows,
                    fmts=[FMT_INT, FMT_INT, FMT_INT, None])
ch4 = BarChart()
ch4.type = "col"
ch4.grouping = "stacked"
ch4.overlap = 100
ch4.title = "基礎的/選択的支出の推移(円/月)※改定年で系列が切り替わる"
ch4.add_data(Reference(ws, min_col=2, min_row=b_start, max_col=3, max_row=b_end), titles_from_data=True)
ch4.set_categories(Reference(ws, min_col=1, min_row=b_start + 1, max_row=b_end))
ch4.height, ch4.width = 7.5, 16
ws.add_chart(ch4, "F%d" % b_start)

put_notes(ws, b_end + 2, [
    "注: 金額はすべて名目値(物価変動を含む)。「増えた=豊かになった」とは限らない。",
    "※1 基礎的支出比率は分類改定ごとに系列が異なるため、2015/2020/2025年は各年の基準系列の値を並べたもので、系列をまたぐ単純比較はできない。",
    "注: 基礎的/選択的支出は2010・2015・2020・2025年の分類改定で接続しない。改定年の前年との増減は出さない設計とした。",
    "注: 二人以上の世帯では実収入・可処分所得・平均消費性向は公表されないため、本ダッシュボードに所得系KPIは置いていない。",
    "注: 2018年1月に調査方法の変更があり、2018年前後の段差は方法変更の影響を含み得る。",
])

# =====================================================================
# シート2: 品目別推移(品目分類)
# =====================================================================
ws2 = wb.create_sheet("2_品目別推移")
ws2.sheet_properties.pageSetUpPr.fitToPage = True
ws2.page_setup.orientation = "landscape"
ws2.page_setup.paperSize = ws2.PAPERSIZE_A4
put_title(
    ws2,
    "日本の家計支出ダッシュボード — 2. 品目別推移",
    "分類: 品目分類(全国) ※シート1・3・4の用途分類とは別体系で、金額は一致しない",
    "出典: 総務省統計局「家計調査(家計収支編)」長期時系列表1-1(月次)・表1-2(年次、月平均は年額÷12の算出値) / " + SRC_DATE,
)


def item_val(year, name, level):
    for r in item_annual:
        if r["year"] == str(year) and r["item_name"] == name and r["level"] == str(level):
            v = fnum(r["amount_yen_per_month_calc"])
            return round(v) if v is not None else None
    return None


ann_rows = []
for name in L1_USE:
    v15, v20, v25 = item_val(2015, name, 1), item_val(2020, name, 1), item_val(2025, name, 1)
    ann_rows.append([name, v15, v20, v25, round(v25 / v15 * 100, 1) if v15 else None])
a_end = write_table(ws2, 5, 1, ["大分類", "2015年", "2020年", "2025年", "指数(2015=100)"],
                    ann_rows, fmts=[None, FMT_INT, FMT_INT, FMT_INT, FMT_PCT1],
                    widths=[20, 13, 13, 13, 15])

idx_years = list(range(2015, 2026))
idx_rows = []
for y in idx_years:
    row = [y]
    for name in L1_USE:
        base = item_val(2015, name, 1)
        v = item_val(y, name, 1)
        row.append(round(v / base * 100, 1) if base and v else None)
    idx_rows.append(row)
i_start = a_end + 2
ws2.cell(row=i_start - 1, column=1, value="大分類10の指数推移(2015年=100、年次・算出値)").font = SUB_FONT
i_end = write_table(ws2, i_start, 1, ["年"] + L1_USE, idx_rows,
                    fmts=[FMT_INT] + [FMT_PCT1] * 10)
ch_idx = LineChart()
ch_idx.title = "大分類10 指数推移(2015=100)"
for col in range(2, 12):
    ch_idx.add_data(Reference(ws2, min_col=col, min_row=i_start, max_row=i_end), titles_from_data=True)
ch_idx.set_categories(Reference(ws2, min_col=1, min_row=i_start + 1, max_row=i_end))
ch_idx.height, ch_idx.width = 8, 18
ws2.add_chart(ch_idx, "A%d" % (i_end + 2))

food_l2 = list(OrderedDict.fromkeys(
    r["item_name"] for r in item_annual
    if r["level"] == "2" and r["category_l1"] == "食料" and r["year"] == "2025"))
f_rows = []
for nm in food_l2:
    v15, v25 = item_val(2015, nm, 2), item_val(2025, nm, 2)
    f_rows.append([nm, v15, v25, round(v25 / v15 * 100, 1) if v15 and v25 else None])
f_start = i_end + 19
ws2.cell(row=f_start - 1, column=1, value="中分類ドリルダウン例: 食料の中分類(円/月)").font = SUB_FONT
f_end = write_table(ws2, f_start, 1, ["中分類(食料)", "2015年", "2025年", "指数(2015=100)"],
                    f_rows, fmts=[None, FMT_INT, FMT_INT, FMT_PCT1])

monthly = {}
for r in item_monthly:
    if r["level"] == "1":
        monthly[(r["year_month"], r["item_name"])] = fnum(r["amount_yen_per_month"])
hm_rows = []
for y in (2024, 2025):
    for m in range(1, 13):
        ym = "%04d-%02d" % (y, m)
        pym = "%04d-%02d" % (y - 1, m)
        row = [ym]
        for name in L1_USE:
            cur, prev = monthly.get((ym, name)), monthly.get((pym, name))
            row.append(round((cur / prev - 1) * 100, 1) if cur and prev else None)
        hm_rows.append(row)
h_start = f_end + 2
ws2.cell(row=h_start - 1, column=1, value="前年同月比ヒートマップ(％、大分類10、2024〜2025年)").font = SUB_FONT
h_end = write_table(ws2, h_start, 1, ["年月"] + L1_USE, hm_rows, fmts=[None] + [FMT_PCT1] * 10)
ws2.conditional_formatting.add(
    "B%d:K%d" % (h_start + 1, h_end),
    ColorScaleRule(start_type="num", start_value=-10, start_color="F8696B",
                   mid_type="num", mid_value=0, mid_color="FFEB84",
                   end_type="num", end_value=10, end_color="63BE7B"))

m_months = sorted(set(r["year_month"] for r in item_monthly))
m_rows = [[ym] + [monthly.get((ym, nm)) for nm in L1_USE] for ym in m_months]
m_start = h_end + 2
ws2.cell(row=m_start - 1, column=1, value="大分類10の月次推移(円/月、2015-01〜2025-12)").font = SUB_FONT
m_end = write_table(ws2, m_start, 1, ["年月"] + L1_USE, m_rows, fmts=[None] + [FMT_INT] * 10)
ch_m = LineChart()
ch_m.title = "大分類10 月次推移(円/月、品目分類)"
for col in range(2, 12):
    ch_m.add_data(Reference(ws2, min_col=col, min_row=m_start, max_row=m_end), titles_from_data=True)
ch_m.set_categories(Reference(ws2, min_col=1, min_row=m_start + 1, max_row=m_end))
ch_m.x_axis.tickLblSkip = 12
ch_m.height, ch_m.width = 9, 20
ws2.add_chart(ch_m, "M%d" % m_start)

put_notes(ws2, m_end + 2, [
    "注: 品目分類と用途分類は別体系であり、大分類名が同じでも金額が一致しない費目がある(例: 2015年の食料は品目分類78,143円 vs 用途分類71,844円)。総額(消費支出)は一致する。シートをまたいだ金額の直接比較はしないこと。",
    "注: 品目分類は2015・2020・2025年に分類改定があり、中分類以下は改定年をまたいで接続しない場合がある。",
    "注: 年次の月平均は年額÷12の算出値。月次表の年平均とは丸めの範囲で差が出ることがある。",
])

# =====================================================================
# シート3: 地域比較(用途分類)
# =====================================================================
ws3 = wb.create_sheet("3_地域比較")
ws3.sheet_properties.pageSetUpPr.fitToPage = True
ws3.page_setup.orientation = "landscape"
ws3.page_setup.paperSize = ws3.PAPERSIZE_A4
put_title(
    ws3,
    "日本の家計支出ダッシュボード — 3. 地域比較",
    "対象: 二人以上の世帯 / 分類: 用途分類(地方10・都市階級4・全国)",
    "出典: 総務省統計局「家計調査(家計収支編)」詳細結果表 第1-1表 / " + SRC_DATE,
)

reg_rows = []
nat25 = region_val(2025, "全国", "消費支出")
for reg in REGIONS:
    v15 = region_val(2015, reg, "消費支出")
    v20 = region_val(2020, reg, "消費支出")
    v25 = region_val(2025, reg, "消費支出")
    reg_rows.append([reg, v15, v20, v25, round(v25 / nat25 * 100, 1)])
reg_rows_sorted = sorted(reg_rows, key=lambda r: -r[3])
r3_end = write_table(ws3, 5, 1, ["地方", "2015年", "2020年", "2025年", "全国=100(2025)"],
                     reg_rows_sorted, fmts=[None, FMT_INT, FMT_INT, FMT_INT, FMT_PCT1],
                     widths=[14, 13, 13, 13, 15])
ch_r = BarChart()
ch_r.type = "bar"
ch_r.title = "地方別 消費支出ランキング(円/月、2025年)"
ch_r.add_data(Reference(ws3, min_col=4, min_row=5, max_row=r3_end), titles_from_data=True)
ch_r.set_categories(Reference(ws3, min_col=1, min_row=6, max_row=r3_end))
ch_r.height, ch_r.width = 8, 14
ch_r.legend = None
ws3.add_chart(ch_r, "G5")

p_start = r3_end + 2
ws3.cell(row=p_start - 1, column=1, value="地方別プロファイル(全国=100、2025年、大分類10)").font = SUB_FONT
p_rows = []
for reg in REGIONS:
    row = [reg]
    for item in L1_USE:
        rv, nv = region_val(2025, reg, item), region_val(2025, "全国", item)
        row.append(round(rv / nv * 100, 1) if rv is not None and nv else None)
    p_rows.append(row)
p_end = write_table(ws3, p_start, 1, ["地方"] + L1_USE, p_rows, fmts=[None] + [FMT_PCT1] * 10)
ws3.conditional_formatting.add(
    "B%d:K%d" % (p_start + 1, p_end),
    ColorScaleRule(start_type="num", start_value=70, start_color="F8696B",
                   mid_type="num", mid_value=100, mid_color="FFFFFF",
                   end_type="num", end_value=130, end_color="638EC6"))

city_rows = []
for cc in CITY_CLASSES:
    tot = region_val(2025, cc, "消費支出")
    city_rows.append([cc] + [round(region_val(2025, cc, it) / tot * 100, 1) for it in L1_USE])
ct_start = p_end + 2
ws3.cell(row=ct_start - 1, column=1, value="都市階級別 費目構成比(％、2025年、算出値)").font = SUB_FONT
ct_end = write_table(ws3, ct_start, 1, ["都市階級"] + L1_USE, city_rows, fmts=[None] + [FMT_PCT1] * 10)
ch_c = BarChart()
ch_c.type = "bar"
ch_c.grouping = "percentStacked"
ch_c.overlap = 100
ch_c.title = "都市階級別 構成比(100％積み上げ、2025年)"
for col in range(2, 12):
    ch_c.add_data(Reference(ws3, min_col=col, min_row=ct_start, max_row=ct_end), titles_from_data=True)
ch_c.set_categories(Reference(ws3, min_col=1, min_row=ct_start + 1, max_row=ct_end))
ch_c.height, ch_c.width = 7, 16
ws3.add_chart(ch_c, "A%d" % (ct_end + 2))

e_rows = [[reg, region_val(2015, reg, "エンゲル係数(％)"), region_val(2025, reg, "エンゲル係数(％)")]
          for reg in REGIONS]
e_start = ct_end + 19
ws3.cell(row=e_start - 1, column=1, value="エンゲル係数(％、地方別、公表値)").font = SUB_FONT
e_end = write_table(ws3, e_start, 1, ["地方", "2015年", "2025年"], e_rows,
                    fmts=[None, FMT_PCT1, FMT_PCT1])
ch_e = BarChart()
ch_e.title = "エンゲル係数(％) 2025年・地方別"
ch_e.add_data(Reference(ws3, min_col=3, min_row=e_start, max_row=e_end), titles_from_data=True)
ch_e.set_categories(Reference(ws3, min_col=1, min_row=e_start + 1, max_row=e_end))
ch_e.height, ch_e.width = 7, 13
ch_e.legend = None
ws3.add_chart(ch_e, "E%d" % e_start)

put_notes(ws3, e_end + 2, [
    "注: 地方値は県庁所在市等の標本に基づく。「その地方の県庁所在市周辺の姿」に近く、都道府県単位の厳密な代表値ではない。",
    "注: 地域・年代の公表は年次(本表では2015・2020・2025年)のみ。月次は全国値のみ公表。",
])

# =====================================================================
# シート4: 世帯主年代別(用途分類)
# =====================================================================
ws4 = wb.create_sheet("4_年代別")
ws4.sheet_properties.pageSetUpPr.fitToPage = True
ws4.page_setup.orientation = "landscape"
ws4.page_setup.paperSize = ws4.PAPERSIZE_A4
put_title(
    ws4,
    "日本の家計支出ダッシュボード — 4. 世帯主年代別",
    "対象: 二人以上の世帯 / 分類: 用途分類(世帯主の年齢階級別)",
    "出典: 総務省統計局「家計調査(家計収支編)」詳細結果表 第3-2表 / " + SRC_DATE,
)

age_rows = [[ag, age_val(2015, ag, "消費支出"), age_val(2020, ag, "消費支出"), age_val(2025, ag, "消費支出")]
            for ag in AGE_GROUPS]
a4_end = write_table(ws4, 5, 1, ["世帯主の年齢階級", "2015年", "2020年", "2025年"], age_rows,
                     fmts=[None, FMT_INT, FMT_INT, FMT_INT], widths=[18, 13, 13, 13])
ch_a = BarChart()
ch_a.title = "年代別 消費支出(円/月、2025年)"
ch_a.add_data(Reference(ws4, min_col=4, min_row=5, max_row=a4_end), titles_from_data=True)
ch_a.set_categories(Reference(ws4, min_col=1, min_row=6, max_row=a4_end))
ch_a.height, ch_a.width = 8, 16
ch_a.legend = None
ws4.add_chart(ch_a, "F5")

hm4_rows = [[ag] + [age_val(2025, ag, it) for it in L1_USE] for ag in AGE_GROUPS]
h4_start = a4_end + 2
ws4.cell(row=h4_start - 1, column=1, value="年代×大分類 金額ヒートマップ(円/月、2025年。色は列内での相対的な大きさ)").font = SUB_FONT
h4_end = write_table(ws4, h4_start, 1, ["年齢階級"] + L1_USE, hm4_rows, fmts=[None] + [FMT_INT] * 10)
for col in range(2, 12):
    col_vals = [r[col - 1] for r in hm4_rows if r[col - 1] is not None]
    ws4.conditional_formatting.add(
        "%s%d:%s%d" % (get_column_letter(col), h4_start + 1, get_column_letter(col), h4_end),
        ColorScaleRule(start_type="num", start_value=min(col_vals), start_color="FFFFFF",
                       end_type="num", end_value=max(col_vals), end_color="2F5597"))

sh_rows = []
for ag in AGE_GROUPS:
    tot = age_val(2025, ag, "消費支出")
    sh_rows.append([ag] + [round(age_val(2025, ag, it) / tot * 100, 1) for it in L1_USE])
s4_start = h4_end + 2
ws4.cell(row=s4_start - 1, column=1, value="年代別 費目構成比(％、2025年、算出値)").font = SUB_FONT
s4_end = write_table(ws4, s4_start, 1, ["年齢階級"] + L1_USE, sh_rows, fmts=[None] + [FMT_PCT1] * 10)
ch_s = BarChart()
ch_s.type = "bar"
ch_s.grouping = "percentStacked"
ch_s.overlap = 100
ch_s.title = "年代別 構成比(100％積み上げ、2025年)"
for col in range(2, 12):
    ch_s.add_data(Reference(ws4, min_col=col, min_row=s4_start, max_row=s4_end), titles_from_data=True)
ch_s.set_categories(Reference(ws4, min_col=1, min_row=s4_start + 1, max_row=s4_end))
ch_s.height, ch_s.width = 9, 17
ws4.add_chart(ch_s, "A%d" % (s4_end + 2))

put_notes(ws4, s4_end + 19, [
    "注: 年代は「世帯主の年齢」であり、世帯構成員全員の年代ではない。",
    "注: 2015年と2025年の同一年代の比較は、同じ世帯の追跡ではなく、別の世帯群の比較(コーホートのずれ)である。",
    "見どころ: 教育費は45～49歳で30,771円/月、85歳以上では204円/月と桁が違う。消費支出のピークは50～54歳(379,853円)。",
])

# =====================================================================
# データシート群(CSVをそのまま投入)
# =====================================================================
DATA_SHEETS = [
    ("data_用途年次", "use_annual.csv"),
    ("data_品目年次", "item_annual.csv"),
    ("data_品目月次", "item_monthly.csv"),
    ("data_基礎選択", "basic_selective_annual.csv"),
    ("data_地域", "region_use.csv"),
    ("data_年代", "age_use.csv"),
    ("data_世帯属性", "household_attributes.csv"),
]
for sheet_name, fname in DATA_SHEETS:
    with open(os.path.join(PROC, fname), newline="", encoding="utf-8") as fh:
        rows = list(csv.reader(fh))
    dws = wb.create_sheet(sheet_name)
    for i, row in enumerate(rows):
        for j, v in enumerate(row):
            c = dws.cell(row=i + 1, column=j + 1)
            if i == 0:
                c.value = v
                c.fill = HDR_FILL
                c.font = HDR_FONT
            else:
                fv = fnum(v)
                c.value = fv if fv is not None else v
                c.font = CELL_FONT
    dws.freeze_panes = "A2"

# =====================================================================
# 指標定義・出典シート
# =====================================================================
wsd = wb.create_sheet("指標定義・出典")
wsd.column_dimensions["A"].width = 22
wsd.column_dimensions["B"].width = 60
wsd.column_dimensions["C"].width = 30
wsd.column_dimensions["D"].width = 52
put_title(wsd, "指標定義・出典・注意点", "対象: 二人以上の世帯", "")
defs = [
    ["消費支出", "1世帯当たり1か月間の消費支出(年平均)", "長期時系列表2/詳細結果表/表1-2", "名目値。物価変動を含む。総額は品目分類・用途分類で一致"],
    ["エンゲル係数", "食料÷消費支出×100(％)。公表値(再掲行)をそのまま使用", "詳細結果表/表2", "用途分類ベース。品目分類の食料では再計算しない"],
    ["基礎的支出・選択的支出", "家計調査の区分(基礎的=生活に不可欠な支出、選択的=余暇・嗜好寄り)", "長期時系列表4", "分類改定(2010・2015・2020・2025年)で接続しない"],
    ["基礎的支出比率", "基礎的÷(基礎的+選択的)×100(算出値)", "表4から算出", "算出値であることを明示。系列をまたぐ増減は出さない"],
    ["構成比", "費目÷消費支出×100(算出値)", "各表から算出", "同一分類内でのみ計算する"],
    ["指数(2015=100)", "対象年金額÷2015年金額×100(算出値)", "品目分類 年次・月次", "分類改定年をまたぐ中分類の指数は注記付き"],
    ["世帯人員・世帯主年齢・持家率", "世帯属性の平均値。金額解釈の補助指標", "詳細結果表", "「1世帯当たり」金額の背景として併記"],
]
d_end = write_table(wsd, 4, 1, ["指標", "定義", "出所", "注意"], defs)
notes = [
    "",
    "【注意点(設計書 §8)】",
    "1. 名目値中心: 物価上昇期は「増えている=豊か」ではない。",
    "2. 分類改定による断絶: 収支項目分類は2010・2015・2020・2025年に改定。基礎的/選択的は改定年で接続しない。品目分類の中分類以下も改定年をまたぐ単純比較は避ける。",
    "3. 品目分類と用途分類は別体系: 大分類名が同じでも金額が一致しない費目がある(例 2015年・食料: 品目78,143円 vs 用途71,844円)。総額(消費支出)は一致する。",
    "4. 調査方法の変更: 2018年1月から家計簿様式・オンライン回答導入等の変更があり、2018年前後の段差は方法変更の影響を含み得る。",
    "5. 標本の性格: 二人以上の世帯の集計世帯数は年平均で約7,300世帯。地方値は県庁所在市等の標本に基づく。",
    "6. 世帯属性の変化: 世帯人員の減少・高齢化が進み、「1世帯当たり」金額の長期比較は世帯構成の変化を含む。1人当たり換算は本サンプルでは行わない。",
    "7. 2025年の年平均値: 取得時点(2026-11-12)で公表済みの値を使用。訂正公表があれば差し替える。",
    "",
    "【出典】",
    "総務省統計局「家計調査(家計収支編)」政府統計コード 00200561(e-Stat)。長期時系列表1-1・1-2・2・4、詳細結果表 第1-1表・第3-2表(2015・2020・2025年)。取得日: 2026-11-12。取得はe-Stat公開ダウンロードのみ(API ID不使用)。",
    "",
    "【更新手順】",
    "scripts/download_sources.py で再取得 → scripts/build_tidy.py で整形 → scripts/build_excel.py で本ファイルを再生成する。データシートの値を手で直さないこと。",
]
put_notes(wsd, d_end + 1, notes)

wb.save(OUT)
print("saved:", OUT)
