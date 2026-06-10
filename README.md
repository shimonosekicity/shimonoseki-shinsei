# 下関市 申請書作成サービス

必要事項をWebフォームに入力するだけで、下関市の各種申請書（Word / PDF）を自動生成するFlaskアプリです。

対応様式（8種類）

| フォーム | 申請書 |
|---|---|
| /teijuu | 定住奨励金支給申請書 |
| /shussan | 出産祝い金支給申請書 |
| /shogai | 障害福祉サービス事業等開始届 |
| /noshin | 農業振興地域整備計画変更申出書 |
| /suido | 排水設備所有者変更届 |
| /kouminkan | 公民館 使用許可申請書・使用料減免申請書・使用中止届 |

## ローカルでの起動

```bash
pip install -r requirements.txt
python app.py
# http://localhost:5000
```

PDF出力には LibreOffice が必要です（無い環境ではPDFボタンが自動的に非表示になり、Word出力のみ利用できます）。

## 構成

- `app.py` — 全申請書を設定駆動で処理。`FORMS` 辞書に1件追加するだけで新しい様式に対応できます
- `templates/base.html` — 全ページ共通レイアウト（各ページは extends で継承）
- `static/style.css` / `static/app.js` — 共通スタイルと共通JS（和暦入力・郵便番号検索など）
- `*_template.docx` — docxtpl 用のWordテンプレート（`{{ 変数名 }}` を差し込み）
- `tools/fix_templates.py` — テンプレートへのプレースホルダ一括追加スクリプト（再実行可）
- `tools/smoke_test.py` — 全ページ表示＋全様式の生成テスト。`python tools/smoke_test.py`

## 新しい申請書の追加手順

1. Wordテンプレートに `{{ 変数名 }}` を配置して リポジトリ直下に置く
2. `app.py` の `FORMS` に1エントリ追加（テンプレート名・ファイル名・申請者名フィールド）
3. `templates/` に入力フォームページを作成（`base.html` を extends、送信先は `/generate/<form_id>`）
4. `app.py` の `PAGES` にURLを追加

## Render へのデプロイ（PDF対応）

PDF変換に LibreOffice を使うため、**Docker ランタイム**でデプロイします。

1. Render のダッシュボードでサービスの Language を `Docker` に変更（または新規作成時に `render.yaml` / `Dockerfile` が自動検出されます）
2. リポジトリを push すれば `Dockerfile`（LibreOffice + 日本語フォント入り）でビルドされます

従来どおり Python ランタイム（`Procfile` + gunicorn）でも動作しますが、その場合PDFボタンは表示されません。
