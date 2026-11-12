#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
サンプル第2弾A「日本の家計支出ダッシュボード」— 公開データ取得スクリプト

e-Stat「家計調査（家計収支編）」(政府統計コード 00200561) の公開ファイルを、
APIアプリケーションIDを使わずに file-download エンドポイントから取得する。
取得日: 2026-11-12（初回取得）。再実行すると data/raw/ に同じファイルを取得し直す。

出典一覧（いずれも総務省統計局「家計調査」）:
  長期時系列表（二人以上の世帯） … stat.go.jp/data/kakei/longtime/ から e-Stat へリンク
    - 表1-1 品目分類（小分類まで）支出金額（月）2000年1月～  statInfId=000040270646 (CSV)
    - 表1-2 品目分類（小分類まで）支出金額（年）2000年～     statInfId=000040270652 (CSV)
    - 表2   用途分類 収入及び支出金額（月・四半期・年・年度） statInfId=000040270657 (Excel)
    - 表4   基礎的支出・選択的支出（二人以上の世帯）          statInfId=000040270661 (Excel)
  詳細結果表（二人以上の世帯・年次）… e-Stat ファイル一覧 cycle=7, tclass3=000000330005
    - 表1-1 用途分類 都市階級・地方・都道府県庁所在市別
        2025年 statInfId=000040409521 (fileKind=4)
        2020年 statInfId=000032048810 (fileKind=4, 実体は旧形式.xls)
        2015年 statInfId=000031369593 (fileKind=0, 旧形式.xls)
    - 表3-2 用途分類 世帯主の年齢階級別
        2025年 statInfId=000040409533 (fileKind=4)
        2020年 statInfId=000032048822 (fileKind=4, 実体は旧形式.xls)
        2015年 statInfId=000031370116 (fileKind=0, 旧形式.xls)
"""
import pathlib
import urllib.request

BASE = "https://www.e-stat.go.jp/stat-search/file-download"
RAW = pathlib.Path(__file__).resolve().parent.parent / "data" / "raw"

FILES = [
    # (statInfId, fileKind, 保存ファイル名)
    ("000040270646", 1, "longtime_1-1_item_monthly_amount.csv"),
    ("000040270652", 1, "longtime_1-2_item_annual_amount.csv"),
    ("000040270657", 4, "longtime_2_use_income_expenditure.xlsx"),
    ("000040270661", 4, "longtime_4_basic_selective.xlsx"),
    ("000040409521", 4, "detail2025_1-1_use_by_region.xlsx"),
    ("000040409533", 4, "detail2025_3-2_use_by_age.xlsx"),
    ("000032048810", 4, "detail2020_1-1_use_by_region.xls"),   # 実体は旧形式 .xls
    ("000032048822", 4, "detail2020_3-2_use_by_age.xls"),     # 実体は旧形式 .xls
    ("000031369593", 0, "detail2015_1-1_use_by_region.xls"),  # 2015年分は fileKind=0 が EXCEL
    ("000031370116", 0, "detail2015_3-2_use_by_age.xls"),
]


def main() -> None:
    RAW.mkdir(parents=True, exist_ok=True)
    for sid, kind, name in FILES:
        url = f"{BASE}?statInfId={sid}&fileKind={kind}"
        dest = RAW / name
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req) as res:
            data = res.read()
        dest.write_bytes(data)
        print(f"{name}: {len(data):,} bytes")


if __name__ == "__main__":
    main()
