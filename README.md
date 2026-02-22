# amazon_scraper

Amazon.co.jp の商品ページ（DPページ）から以下4項目を取得するスクレイパー。

- **この商品について** (feature-bullets)
- **メーカーによる説明** (A+ content)
- **商品情報** (Technical details)
- **商品の説明** (Product description)

## セットアップ

```bash
pip install -r requirements.txt
```

## 使い方

### 1. 入力ファイルの準備

`Input/asin_list.xlsx` を作成します（サンプル生成スクリプトあり）:

```bash
python create_sample_input.py
```

Excel形式:
| A列 (ASIN) |
|-------------|
| ASIN        |  ← 1行目: ヘッダー
| B09DX1R4RQ  |  ← 2行目以降: ASINコード
| B0XXXXXXXX  |

### 2. スクレイパー実行

```bash
python amazon_scraper.py
```

### 3. 結果確認

`Output/asin_results.xlsx` に結果が出力されます:

| A列 (ASIN) | B列 (この商品について) | C列 (メーカーによる説明) | D列 (商品情報) | E列 (商品の説明) |
|-------------|----------------------|------------------------|---------------|-----------------|
| B09DX1R4RQ  | ・特徴1...           | メーカーテキスト...      | 項目: 値...    | 説明テキスト...   |

## 注意事項

- Amazonのレート制限を考慮し、リクエスト間に2〜5秒のランダムな待機時間を設けています
- 大量のASINを処理する場合は間隔を長めに設定してください
- Amazonのページ構造が変更された場合、セレクターの調整が必要になることがあります
