# 進路事務

## 私立高校 Web出願チェックシート（長崎県・中学3年生用）

| ファイル | 用途 |
|---|---|
| `output/私立高校_Web出願チェックシート.docx` | 生徒に配る印刷用（A4・1枚、生徒欄と先生欄つき） |
| `output/私立高校_Web出願_進捗管理.xlsx` | 担任がクラス全員の進み具合を一覧で管理する表 |
| `output/私立高校_Web出願チェック.html` | 画面上でチェックできる Web 版 |

チェック項目は `data/checklist.json` にまとめてあります。内容を変えたら、次のコマンドで3つの形式を作り直します。

```bash
node scripts/gansho_checklist.js        # Word
python3 scripts/gansho_progress_xlsx.py # Excel
python3 scripts/build_web.py            # Web
```

## 副申書 必要条件一覧

| ファイル | 用途 |
|---|---|
| `output/副申書_必要条件一覧.xlsx` | 高校ごとの副申書の要否と基準（３年間／各学年／３年時の欠席日数・評定不能・健康上の配慮）を、各高校の担当が入力する一覧 |

先頭行は本校（戸町中）の基準です。各学年で欠席20日以上、成績に（　）表記や評定不能がある、試験で配慮が必要、のいずれかに当てはまる生徒に副申書をつけます。高校名・担当・本校基準は `data/fukushinsho_criteria.json` にまとめてあり、次のコマンドで作り直します。

```bash
python3 scripts/fukushinsho_criteria_xlsx.py
```
