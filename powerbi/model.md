# Power BI データモデル定義 — 日本の家計支出ダッシュボード

- 対象: 総務省統計局「家計調査(家計収支編)」(政府統計コード 00200561)、二人以上の世帯
- データ: `data/processed/` の整形済みCSV 7点(取得日 2026-11-12、e-Stat公開ダウンロード)
- 設計方針: スター・スキーマ。ファクトは分類体系・粒度ごとに分離し、**用途分類と品目分類を混在させない**。
  Power BI Desktop での再現手順は本書末尾。

---

## 1. テーブル一覧

### 1.1 ファクト(6)

| テーブル | 元CSV | 粒度 | 備考 |
|---|---|---|---|
| FactUseAnnual | use_annual.csv | 年 × 用途分類項目(全137項目) | ページ1の基礎。金額・係数・人員等の項目を item_name で区別 |
| FactItemAnnual | item_annual.csv | 年 × 品目分類(大分類10+中分類50+総額) | ページ2の年次。amount_yen_per_month_calc(年額÷12)を使う |
| FactItemMonthly | item_monthly.csv | 年月 × 品目分類(大・中) | ページ2の月次(2015-01〜2025-12) |
| FactRegion | region_use.csv | 年(2015/2020/2025) × 地域(全国・都市階級4・地方10) × 用途分類項目(18) | ページ3 |
| FactAge | age_use.csv | 年(同上) × 世帯主年齢階級(平均+12階級) × 用途分類項目(18) | ページ4 |
| FactBasicSelective | basic_selective_annual.csv | 年 × 分類系列 | ページ1。改定年は2系列が併存する |

### 1.2 ディメンション(6)

| テーブル | 生成元 | 内容 |
|---|---|---|
| DimYear | 2000〜2025を生成 | 年キー。年次ファクト専用(月次とは分ける) |
| DimDate | 2000-01〜2026-09を生成 | 年月キー(YYYYMM)・年・月。月次ファクト専用 |
| DimItemUse | use_annual.csv の項目一覧 | 用途分類コード・名称・階層レベル(コードのドット数)・表示順 |
| DimItemProduct | item_annual.csv の項目一覧 | 品目コード・大分類・中分類・階層レベル |
| DimRegion | region_use.csv | 区分(全国/都市階級/地方)・地域名・表示順 |
| DimAge | age_use.csv | 年齢階級・表示順。「平均」は IsAverage フラグで別扱い |

### 1.3 補助(1)

| テーブル | 生成元 | 内容 |
|---|---|---|
| DimAttributes | household_attributes.csv | 世帯属性(世帯人員・世帯主の年齢・持家率など)。KPIカード用に「全国」「平均」だけを参照 |

---

## 2. 列定義(主要ファクト)

### FactUseAnnual
| 列 | 型 | 説明 |
|---|---|---|
| YearKey | 整数 | 年(DimYear と接続) |
| ItemKey | テキスト | 用途分類コード(item_code。空欄の再掲・属性行は item_name から生成したキー) |
| ItemName | テキスト | 項目名 |
| Amount | 10進数 | 支出金額(円/月)。エンゲル係数等の％項目も同じ列に入るため、必ず項目で絞って使うこと |
| Source | テキスト | 出典表名 |

### FactItemAnnual / FactItemMonthly
| 列 | 型 | 説明 |
|---|---|---|
| YearKey / YearMonthKey | 整数 | 年 / 年月(YYYYMM) |
| ItemKey | テキスト | 品目コード |
| CategoryL1 / CategoryL2 | テキスト | 大分類 / 中分類 |
| Amount | 10進数 | 円/月(年次は年額÷12の算出値) |

### FactRegion / FactAge
| 列 | 型 | 説明 |
|---|---|---|
| YearKey | 整数 | 2015・2020・2025 のみ |
| RegionKey / AgeKey | テキスト | 地域名 / 年齢階級 |
| ItemName | テキスト | 用途分類項目(消費支出・大分類10・再掲6・エンゲル係数) |
| Amount | 10進数 | 円/月(エンゲル係数は％値) |

### FactBasicSelective
| 列 | 型 | 説明 |
|---|---|---|
| YearKey | 整数 | 年 |
| ClassificationSeries | テキスト | 分類系列(2010年改定前/2010年改定基準/2015年改定基準/2020年改定基準/2025年改定基準) |
| BasicAmount / SelectiveAmount | 10進数 | 基礎的支出・選択的支出(円/月) |
| BasicSharePct | 10進数 | 基礎的支出比率(算出値、CSVの basic_share_pct_calc) |

---

## 3. リレーション

すべて単方向(ディメンション → ファクト、1対多):

```
DimYear  ──< FactUseAnnual / FactItemAnnual / FactRegion / FactAge / FactBasicSelective
DimDate  ──< FactItemMonthly
DimItemUse    ──< FactUseAnnual
DimItemProduct──< FactItemAnnual / FactItemMonthly
DimRegion     ──< FactRegion
DimAge        ──< FactAge
```

方針:
- **双方向フィルターは使わない。** スライサーの効きはディメンション経由に限定する。
- 年次と月次で日付ディメンションを分ける(DimYear / DimDate)。不連続な日付テーブルに対する
  DATEADD 等の時間インテリジェンスは使わず、前年比は年キー・年月キーの算術で実装する(measures.dax 参照)。
- 用途分類コードと品目コードは見かけが同じでも別物。共通ディメンションに統合しない。
- 地域・年齢ファクトは 2015・2020・2025 の3時点のみ。ページ1・2の年スライサーと共有しない
  (レポートレベルではなくページレベルのフィルターで年を切り替える)。

---

## 4. 取り込み手順(Power Query)

1. 「データを取得 > テキスト/CSV」で `data/processed/*.csv` を7点読み込む(UTF-8)。
2. 上表のとおり列名を付け替え、YearKey 等のキーを追加する。
   - use_annual.csv の item_code が空欄の行(再掲・属性項目)は、item_name をそのまま ItemKey とする。
3. ディメンション6点は「参照」クエリ+重複削除で生成する(DimYear・DimDate はカレンダー生成)。
4. モデルビューで §3 のリレーションを作成し、ファクト側のキー列は非表示にする。
5. measures.dax のメジャーを各ファクト用メジャーテーブルに登録する。
6. pages.md の仕様どおりに4ページを配置する。

## 5. 注意(モデルの前提)

- FactUseAnnual の Amount には金額・％・人員が混在する(元表2の仕様)。メジャーは必ず項目名で
  フィルターすること。ビジュアル側での列の直接集計(SUM の暗黙メジャー)は禁止とする。
- FactBasicSelective は改定年に2系列が重なる。年でそのまま合計すると二重計上になるため、
  描画時は必ず1系列に絞る(既定は各年の最新系列)。系列をまたぐ増減メジャーは意図的に作らない。
- エンゲル係数は公表値を項目として保持し、再計算しない(品目分類の食料で割り直さない)。
