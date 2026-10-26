# 進路事務

## 私立高校 Web出願チェックシート（長崎県・中学3年生用）

| ファイル | 用途 |
|---|---|
| `output/私立高校_Web出願チェックシート.docx` | 生徒に配る印刷用（A4・2ページ、生徒欄と担任欄つき） |
| `output/私立高校_Web出願_進捗管理.xlsx` | 担任がクラス全員の進み具合を一覧で管理する表 |
| `output/私立高校_Web出願チェック.html` | 画面上でチェックできる Web 版 |

チェック項目は `data/checklist.json` にまとめてあります。内容を変えたら、次のコマンドで3つの形式を作り直します。

```bash
node scripts/gansho_checklist.js        # Word
python3 scripts/gansho_progress_xlsx.py # Excel
python3 scripts/build_web.py            # Web
```
