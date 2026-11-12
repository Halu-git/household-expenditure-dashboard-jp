#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
サンプル第2弾A「日本の家計支出ダッシュボード」— 整形スクリプト

data/raw/ の e-Stat「家計調査」公開ファイルを読み、data/processed/ に
タイidy形式（ロング形式）の分析用CSVを出力する。再現可能：生ファイルさえあれば
このスクリプト1本で全整形済みCSVを再生成できる。

前提: Python 3 + openpyxl。旧形式 .xls の変換に LibreOffice (soffice) を使う
     （xlrd に依存しない）。
出力:
  item_annual.csv             品目分類・年次（大分類/中分類、2000-2025）
  item_monthly.csv            品目分類・月次（大分類/中分類、2015-01～2025-12）
  use_annual.csv              用途分類・年次（長期時系列表2、全項目、2000-2025）
  basic_selective_annual.csv  基礎的支出・選択的支出（2005-2025、分類改定の系列付き）
  region_use.csv              用途分類×地域（2015/2020/2025、全国・都市階級・地方）
  age_use.csv                 用途分類×世帯主年齢階級（2015/2020/2025）
  household_attributes.csv    世帯属性（世帯人員・世帯主年齢・持家率など）
金額はいずれも「1世帯当たり1か月間・年平均」円。'…'・'-' は欠損として空欄にする。
"""
import csv
import pathlib
import re
import subprocess
import sys
import tempfile

import openpyxl

ROOT = pathlib.Path(__file__).resolve().parent.parent
RAW = ROOT / "data" / "raw"
OUT = ROOT / "data" / "processed"
OUT.mkdir(parents=True, exist_ok=True)

MISSING = {"", "...", "…", "-", "***", "－"}


def to_num(v):
    """e-Statの表記ゆれを数値化する。欠損は None。"""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).replace(",", "").replace("，", "").strip()
    if s in MISSING:
        return None
    try:
        return float(s)
    except ValueError:
        return None


def norm_label(s):
    """全角スペース等を取り除いて見出しを正規化する。"""
    if s is None:
        return ""
    return re.sub(r"[\s　]+", "", str(s))


def write_csv(name, header, rows):
    path = OUT / name
    with open(path, "w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(header)
        w.writerows(rows)
    print(f"{name}: {len(rows):,} rows")
    return rows


# ---------------------------------------------------------------- 品目分類CSV
def parse_item_csv(path):
    """長期時系列表1-1(月)/1-2(年) のCSV(Shift_JIS, ワイド形式)をパースする。"""
    with open(path, encoding="cp932", newline="") as f:
        rows = list(csv.reader(f))
    # '表側連番' を含む行が項目ヘッダ行
    hidx = next(i for i, r in enumerate(rows) if "表側連番" in r)
    header = rows[hidx]
    col = {name: header.index(name) for name in ("階層", "大分類", "中分類", "小分類", "品目分類")}
    # 期間ラベル: 年次はヘッダ行の1行上に '2000年' が並ぶ。
    # 月次は2行上に年が疎に（各年1月列のみ）入り、ヘッダ行に '1月' 形式で月が並ぶ。
    period_cols = {}  # 列index -> ('Y', year) or ('M', (year, month))
    for j, cell in enumerate(rows[hidx - 1]):
        m = re.fullmatch(r"(\d{4})年(\d{1,2})月", (cell or "").strip())
        if m:
            period_cols[j] = ("M", (int(m.group(1)), int(m.group(2))))
            continue
        m = re.fullmatch(r"(\d{4})年", (cell or "").strip())
        if m:
            period_cols[j] = ("Y", int(m.group(1)))
    if not any(k == "M" for k, _ in period_cols.values()):
        year_at, cur = {}, None
        for j, cell in enumerate(rows[hidx - 2]):
            m = re.fullmatch(r"(\d{4})年", (cell or "").strip())
            if m:
                cur = int(m.group(1))
            year_at[j] = cur
        for j, cell in enumerate(header):
            m = re.fullmatch(r"(\d{1,2})月", (cell or "").strip())
            if m and year_at.get(j):
                period_cols[j] = ("M", (year_at[j], int(m.group(1))))
    # 階層: 1=消費支出(総額), 2=大分類, 3=中分類, 4=小分類。
    # 大分類/中分類列はコードで、名称は品目分類列にある。行は階層順に並ぶ。
    items = []  # (level, item_code, 大分類名, 中分類名, 品目名, {列index: 値})
    in_items = False
    cur_l1 = cur_l2 = ""
    for r in rows[hidx + 1:]:
        if len(r) <= max(col.values()):
            continue
        name = (r[col["品目分類"]] or "").strip()
        if not name:
            continue
        if name == "消費支出":
            in_items = True
        if not in_items:
            continue  # 世帯属性ブロックは対象外（属性は詳細結果表・表2から取る）
        tier = (r[col["階層"]] or "").strip()
        c1 = (r[col["大分類"]] or "").strip()
        c2 = (r[col["中分類"]] or "").strip()
        vals = {j: to_num(r[j]) if j < len(r) else None for j in period_cols}
        if tier == "1":
            items.append((0, "", "消費支出", "", name, vals))
            cur_l1 = cur_l2 = ""
        elif tier == "2":
            cur_l1, cur_l2 = name, ""
            items.append((1, c1, name, "", name, vals))
        elif tier == "3":
            cur_l2 = name
            items.append((2, f"{c1}.{c2}", cur_l1, name, name, vals))
        # tier == '4'（小分類）はダッシュボード対象外
    return period_cols, items


def build_item_tables():
    # 年次（表1-2の金額は「1世帯当たり年間支出金額」。月平均は÷12の算出値を併記）
    pcols, items = parse_item_csv(RAW / "longtime_1-2_item_annual_amount.csv")
    rows = []
    for level, code, big, mid, name, vals in items:
        for j, (kind, year) in pcols.items():
            v = vals.get(j)
            if v is None:
                continue
            rows.append([year, level, code, big, mid, name,
                         int(v) if v == int(v) else v, round(v / 12, 1),
                         "家計調査 長期時系列表1-2 品目分類(年)"])
    write_csv("item_annual.csv",
              ["year", "level", "item_code", "category_l1", "category_l2", "item_name",
               "amount_yen_per_year", "amount_yen_per_month_calc", "source"], rows)
    # 月次（設計対象期間 2015-01～2025-12 に限定。生データは2000-01～2026-09）
    pcols, items = parse_item_csv(RAW / "longtime_1-1_item_monthly_amount.csv")
    rows = []
    for level, code, big, mid, name, vals in items:
        for j, (kind, (y, mo)) in pcols.items():
            if not (2015 <= y <= 2025):
                continue
            v = vals.get(j)
            if v is None:
                continue
            rows.append([f"{y:04d}-{mo:02d}", y, mo, level, code, big, mid, name,
                         int(v) if v == int(v) else v,
                         "家計調査 長期時系列表1-1 品目分類(月)"])
    write_csv("item_monthly.csv",
              ["year_month", "year", "month", "level", "item_code", "category_l1",
               "category_l2", "item_name", "amount_yen_per_month", "source"], rows)


# ---------------------------------------------------------------- 用途分類・年次（表2）
def era_to_year(label):
    s = norm_label(label)
    m = re.search(r"平成(\d+)年", s)
    if m:
        return 1988 + int(m.group(1))
    m = re.search(r"令和(\d+)年", s)
    if m:
        return 2018 + int(m.group(1))
    if "令和元年" in s:
        return 2019
    return None


def row_name(cells, idxs):
    for i in idxs:
        if i < len(cells) and cells[i] is not None and str(cells[i]).strip():
            return str(cells[i]).strip()
    return ""


def value_kind(name):
    if "（％）" in name or "（%)" in name:
        return "percent"
    if "（人）" in name:
        return "persons"
    if "（歳）" in name:
        return "years_old"
    return "yen_per_month"


def build_use_annual():
    wb = openpyxl.load_workbook(RAW / "longtime_2_use_income_expenditure.xlsx",
                                read_only=True, data_only=True)
    ws = wb["支出金額（年）"]
    grid = [list(r) for r in ws.iter_rows(values_only=True)]
    wb.close()
    # 見出し行(11行目=index10)から年次列を特定
    year_cols = {}
    for j in range(15, len(grid[10])):
        y = era_to_year(grid[10][j])
        if y:
            year_cols[j] = y
    rows = []
    for r in grid:
        if not r or r[0] != "10101":
            continue
        code = r[7] if len(r) > 7 else None
        code = str(code).strip() if code and re.fullmatch(r"\d+(\.\d+)*", str(code).strip()) else ""
        name = row_name(r, [10, 11, 12, 13, 8, 9, 7])
        if not name:
            continue
        for j, y in year_cols.items():
            v = to_num(r[j]) if j < len(r) else None
            if v is None:
                continue
            rows.append([y, code, name, value_kind(name),
                         int(v) if v == int(v) else v,
                         "家計調査 長期時系列表2 用途分類 支出金額(年)"])
    write_csv("use_annual.csv",
              ["year", "item_code", "item_name", "value_kind", "value", "source"], rows)


# ---------------------------------------------------------------- 表4 基礎的/選択的
def build_basic_selective():
    wb = openpyxl.load_workbook(RAW / "longtime_4_basic_selective.xlsx",
                                read_only=True, data_only=True)
    ws = wb["年"]
    grid = [list(r) for r in ws.iter_rows(values_only=True)]
    wb.close()
    series_labels = ["2010年改定前", "2010年改定基準", "2015年改定基準",
                     "2020年改定基準", "2025年改定基準"]
    rows = []
    seg = 0
    prev_year = None
    for r in grid:
        if len(r) < 9 or not r[1]:
            continue
        m = re.search(r"\((\d{4})年\)", str(r[1]))
        if not m:
            continue
        year = int(m.group(1))
        if prev_year is not None and year <= prev_year:
            seg += 1  # 同一年の重複行 = 分類改定による新系列の開始
        prev_year = year
        basic, selective = to_num(r[3]), to_num(r[4])
        if basic is None or selective is None:
            continue
        share = round(basic / (basic + selective) * 100, 2)
        rows.append([year, series_labels[seg], int(basic), int(selective), share,
                     to_num(r[5]), to_num(r[6]), to_num(r[7]), to_num(r[8]),
                     "家計調査 長期時系列表4 基礎的支出・選択的支出(年)"])
    write_csv("basic_selective_annual.csv",
              ["year", "classification_series", "basic_yen_per_month",
               "selective_yen_per_month", "basic_share_pct_calc",
               "basic_nominal_growth_pct", "selective_nominal_growth_pct",
               "basic_real_growth_pct", "selective_real_growth_pct", "source"], rows)


# ---------------------------------------------------------------- 詳細結果表（地域・年代）
BIG10 = {"食料", "住居", "光熱・水道", "家具・家事用品", "被服及び履物",
         "保健医療", "交通・通信", "教育", "教養娯楽", "その他の消費支出"}
REGION_LABELS = ["全国", "大都市", "中都市", "小都市A", "小都市B・町村",
                 "北海道", "東北", "関東", "北陸", "東海", "近畿", "中国", "四国", "九州", "沖縄"]
# 年代見出しは中間階級に「歳」が付かない（'35～39'）ため、正規ラベルへ付け替える
AGE_LABELS = ["平均", "～34歳", "35～39", "40～44", "45～49", "50～54",
              "55～59", "60～64", "65～69", "70～74", "75～79",
              "80～84", "85歳～"]
AGE_CANON = {f"{a}～{b}": f"{a}～{b}歳" for a, b in
             [(35, 39), (40, 44), (45, 49), (50, 54), (55, 59),
              (60, 64), (65, 69), (70, 74), (75, 79), (80, 84)]}
ATTR_KEYS = ("世帯数分布", "集計世帯数", "世帯人員", "世帯主の年齢", "持家率")


def convert_xls(path, tmpdir):
    subprocess.run(["soffice", "--headless", "--convert-to", "xlsx",
                    "--outdir", tmpdir, str(path)],
                   check=True, capture_output=True)
    return pathlib.Path(tmpdir) / (path.stem + ".xlsx")


def parse_detail(path, sheet, year, labels, tmpdir, canon=None):
    if path.suffix == ".xls":
        path = convert_xls(path, tmpdir)
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[sheet]
    grid = [list(r) for r in ws.iter_rows(values_only=True)]
    wb.close()
    # 見出し行: 14列目(idx13)に '用 途 分 類'（空白混じり）を含む行
    hidx = next(i for i, r in enumerate(grid)
                if len(r) > 13 and r[13] and "用途分類" in norm_label(r[13]))
    header = grid[hidx]
    cont = grid[hidx + 1] if hidx + 1 < len(grid) else []
    colmap = {}  # 列index -> ラベル
    for j in range(15, len(header)):
        lab = norm_label(header[j])
        if j < len(cont) and cont[j] and norm_label(cont[j]):
            lab += norm_label(cont[j])
        lab = lab.replace("小都市Ａ", "小都市A").replace("小都市Ｂ", "小都市B")
        if lab in labels:
            colmap[j] = (canon or {}).get(lab, lab)
    items, attrs = [], []
    for r in grid[hidx + 1:]:
        seqv = to_num(r[5]) if len(r) > 5 else None
        if len(r) < 16 or seqv is None:
            continue  # 連番のない行（注記など）は対象外
        seq = int(seqv)
        code = r[7] if len(r) > 7 else None
        code = str(code).strip() if code and re.fullmatch(r"\d+(\.\d+)*", str(code).strip()) else ""
        name = row_name(r, [10, 11, 12, 13, 8, 9, 7])
        if not name:
            continue
        vals = {lab: (to_num(r[j]) if j < len(r) else None) for j, lab in colmap.items()}
        if seq <= 15 or (not code and any(name.startswith(k) for k in ATTR_KEYS)):
            if any(name.startswith(k) for k in ATTR_KEYS):
                attrs.append((name, vals))
            continue
        # 大分類は分類コードが1桁の行だけを採る（交際費内訳の同名小分類を除外）
        keep = (name == "消費支出"
                or (name in BIG10 and re.fullmatch(r"\d+", code or "") is not None)
                or name.replace("（", "(").startswith("(再掲")
                or name.startswith("エンゲル係数"))
        if keep:
            items.append((code, name, vals))
    return items, attrs


def build_detail_tables(tmpdir):
    region_rows, age_rows, attr_rows = [], [], []
    region_files = [(2015, "detail2015_1-1_use_by_region.xls", "1-1二人以上の世帯"),
                    (2020, "detail2020_1-1_use_by_region.xls", "二人"),
                    (2025, "detail2025_1-1_use_by_region.xlsx", "二人")]
    age_files = [(2015, "detail2015_3-2_use_by_age.xls", "3-2全国・二人以上の世帯"),
                 (2020, "detail2020_3-2_use_by_age.xls", "二人"),
                 (2025, "detail2025_3-2_use_by_age.xlsx", "二人")]
    for year, fname, sheet in region_files:
        items, attrs = parse_detail(RAW / fname, sheet, year, REGION_LABELS, tmpdir)
        for code, name, vals in items:
            for lab, v in vals.items():
                if v is None:
                    continue
                grp = "全国" if lab == "全国" else ("都市階級" if lab in REGION_LABELS[1:5] else "地方")
                region_rows.append([year, grp, lab, code, name,
                                    int(v) if v == int(v) else v,
                                    "家計調査 詳細結果表 第1-1表 用途分類 都市階級・地方別"])
        for name, vals in attrs:
            for lab, v in vals.items():
                if v is None:
                    continue
                attr_rows.append([year, "地域", lab, name,
                                  int(v) if v == int(v) else v,
                                  "家計調査 詳細結果表 第1-1表"])
    for year, fname, sheet in age_files:
        items, attrs = parse_detail(RAW / fname, sheet, year, AGE_LABELS, tmpdir,
                                    canon=AGE_CANON)
        for code, name, vals in items:
            for lab, v in vals.items():
                if v is None:
                    continue
                age_rows.append([year, lab, code, name,
                                 int(v) if v == int(v) else v,
                                 "家計調査 詳細結果表 第3-2表 用途分類 世帯主の年齢階級別"])
        for name, vals in attrs:
            for lab, v in vals.items():
                if v is None:
                    continue
                attr_rows.append([year, "世帯主年齢階級", lab, name,
                                  int(v) if v == int(v) else v,
                                  "家計調査 詳細結果表 第3-2表"])
    write_csv("region_use.csv",
              ["year", "region_group", "region", "item_code", "item_name",
               "amount_yen_per_month", "source"], region_rows)
    write_csv("age_use.csv",
              ["year", "age_group", "item_code", "item_name",
               "amount_yen_per_month", "source"], age_rows)
    write_csv("household_attributes.csv",
              ["year", "dimension", "group_name", "attribute", "value", "source"],
              attr_rows)


def main():
    build_item_tables()
    build_use_annual()
    build_basic_selective()
    with tempfile.TemporaryDirectory() as tmpdir:
        build_detail_tables(tmpdir)
    print("done.")


if __name__ == "__main__":
    sys.exit(main())
