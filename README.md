# 日本の家計支出ダッシュボード — 家計調査(二人以上の世帯)

総務省統計局「家計調査(家計収支編)」の公開データだけで、「家計の財布は今どこに向かっているか」を読むダッシュボードを作ったサンプルプロジェクトです。
同じ整形済みソースデータから、**Excel完成品**と**Power BI設計パック**の2つの見せ方を用意しています。

## 主なファクト(検証済み)

- 消費支出(1世帯当たり月額)は 2015年 287,373円 → 2020年 277,926円(コロナ禍で減少)→ 2025年 314,001円。独立した4つの表で完全一致を確認済み
- エンゲル係数は 2015年 25.0% → 2025年 28.6%
- 2025年の地方別では関東 337,946円 に対し沖縄 235,742円。世帯主年代別では50〜54歳 379,853円がピーク
- 世帯人員は 3.02人(2015年)→ 2.87人(2025年)、世帯主の平均年齢は 58.8歳 → 60.7歳

## 成果物

- `excel/household_expenditure_dashboard.xlsx`:Excelダッシュボード完成品(概要/品目推移/地域/年代の4シート+データ7シート+指標定義・出典シート、ネイティブグラフ付き)
- `powerbi/model.md`:データモデル定義(スター・スキーマ、テーブル・リレーション・取り込み手順)
- `powerbi/measures.dax`:DAXメジャー一式(メジャーブランチング方針)
- `powerbi/pages.md`:ページ配置仕様と公開前チェックリスト
- `data/processed/`:整形済みCSV 7点(tidy形式)
- `design.md`:設計書(データ基盤・検証結果・指標定義・注意点)

## 使い方(再現手順)

1. `scripts/download_sources.py` で e-Stat の公開ファイルを取得する(API IDは不要、公開ダウンロードのみ)
2. `scripts/build_tidy.py` で整形済みCSVを再生成する
3. `scripts/build_excel.py` で Excel ダッシュボードを再生成する
4. Power BI は `powerbi/` の定義どおりに Power BI Desktop でモデルを組み、DAXを登録する

## データ出典

- 総務省統計局「家計調査(家計収支編)」政府統計コード 00200561(e-Stat)
  - 長期時系列表1-1・1-2・2・4(二人以上の世帯)
  - 詳細結果表 第1-1表(都市階級・地方別)、第3-2表(世帯主の年齢階級別)。2015・2020・2025年
  - 取得日: 2026-11-12

各データの利用条件・出典表示は、それぞれの提供元の規約に従います。生データ(raw)は容量の都合で同梱せず、取得スクリプトでの再現を正としています。

## 読むときの注意(重要)

- **品目分類と用途分類は別体系**です。大分類名が同じでも金額が一致しない費目があります(例: 2015年の食料は品目分類 78,143円 vs 用途分類 71,844円)。総額(消費支出)は一致します。本リポジトリの成果物はページ・シートごとに分類を固定し、出典を明記しています
- 金額は名目値です。物価上昇を含むため「増えた=豊か」とは限りません
- 基礎的/選択的支出は分類改定(2010・2015・2020・2025年)で系列が接続しません。改定年をまたぐ増減は算出していません
- 二人以上の世帯では実収入・可処分所得は公表されないため、所得系の指標は置いていません
- 地方値は県庁所在市等の標本に基づき、都道府県単位の厳密な代表値ではありません

## 開示

集計・可視化のコードはAI支援を用いて作成しました。データ出典の確認、集計条件の設定、結果の解釈と結論の責任は筆者が負っています。

---

## Summary (English)

A household-expenditure dashboard built solely from Japan's official Family Income and Expenditure Survey (Statistics Bureau, MIC; two-or-more-person households), downloaded from e-Stat without an API key. Deliverables: a finished Excel dashboard (4 dashboard sheets with native charts), a Power BI design pack (star-schema model definition, DAX measures, page specs), and tidy CSVs with reproducible build scripts. Key verified facts: monthly consumption expenditure per household was JPY 287,373 (2015), 277,926 (2020) and 314,001 (2025) — cross-checked across four independent published tables; the Engel coefficient rose from 25.0% to 28.6%. Caution: the survey's two classification systems (by item / by purpose) are not interchangeable — category totals differ even though overall consumption matches.
