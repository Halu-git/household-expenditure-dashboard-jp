# PBIX組み立てメモ（サンプル2号A）

設計書（design.md／model.md／measures.dax）は正規化した列名（YearKey・Amount等）を前提にした「設計仕様」。実際のPBIX組み立てでは、Kaiの作業量を最小化するためCSVの元列名のまま組み、DAXは実測版 `measures_build.dax` を使う。意味・方針は正規版と同一。

## 列名の対応（正規 → 実測）
- FactUseAnnual[Amount] → [value]／DimItemUse[ItemName] → [item_name]
- DimYear[YearKey] → DimYear[year]
- FactItemAnnual[Amount] → [amount_yen_per_month_calc]（年額÷12の月額換算）
- FactItemMonthly[Amount] → [amount_yen_per_month]
- DimDate[YearMonthKey] → DimDateは year_month（"YYYY-MM"文字列）をキーにし、前年同月は year・month の算術で実装
- FactRegion[Amount] → [amount_yen_per_month]／DimRegion[Region] → [region]
- FactAge[Amount] → [amount_yen_per_month]
- FactBasicSelective[BasicAmount] → [basic_yen_per_month]／[SelectiveAmount] → [selective_yen_per_month]
- DimAttributes（household_attributes を改名）[Value] → [value]、[Dimension] → [dimension]、[GroupName] → [group_name]、[Attribute] → [attribute]

## キーの実測上の注意（CSVを実測して確認済み）
- item_code は総額行等で空欄がある（item_annual 26行、item_monthly 132行、use_annual 536行）ため、ディメンションのキーは item_name（全61品目／137項目で一意）を使う
- use_annual の value にはエンゲル係数（％）の行も同居する。項目で絞る設計は正規版どおり。value_kind は全行 yen_per_month で区別には使えない
- household_attributes の全国・地域・2025年：世帯人員 2.87、世帯主の平均年齢 60.7（メジャーの参照先として実在確認済み）

## リレーション（すべて単方向・ディメンション→ファクト）
- DimYear[year] → FactUseAnnual / FactItemAnnual / FactRegion / FactAge / FactBasicSelective の [year]
- DimDate[year_month] → FactItemMonthly[year_month]
- DimItemUse[item_name] → FactUseAnnual[item_name]
- DimItemProduct[item_name] → FactItemAnnual[item_name]、FactItemMonthly[item_name]
- DimRegion[region] → FactRegion[region]
- DimAge[age_group] → FactAge[age_group]
- DimAttributes にはリレーションを張らない（KPIカード用にメジャー内で直接絞る）

## ディメンションの作り方（Power Query・参照クエリ）
- DimYear ← FactUseAnnual の参照：year のみ残して重複削除
- DimItemUse ← FactUseAnnual の参照：item_code・item_name を残し、item_name で重複削除
- DimItemProduct ← FactItemAnnual の参照：item_code・item_name・category_l1・category_l2 を残し、item_name で重複削除
- DimDate ← FactItemMonthly の参照：year_month・year・month を残して重複削除
- DimRegion ← FactRegion の参照：region・region_group を残し、region で重複削除
- DimAge ← FactAge の参照：age_group のみ残して重複削除
