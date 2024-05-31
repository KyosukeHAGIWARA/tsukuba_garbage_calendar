# Tsukuba Garbage Calendar
茨城県つくば市のごみ収集カレンダーを扱いやすい JSON ファイルに変換・整形するスクリプト

## What I Want
つくば市はオープンデータとして、つくば市内のごみ収集カレンダーを年度ごとに公式 HP に公開しています。  
ref: [ごみ収集カレンダー／つくば市公式ウェブサイト](https://www.city.tsukuba.lg.jp/soshikikarasagasu/seikatsukankyobukankyoeiseika/gyomuannai/2/1000820.html)

しかし、公開されているのは、PDF 形式のデータか月ごとの Excel ファイルのみで、CSV や JSON データが存在しないのが不便でした。

先例で2017年頃に PDF からデータを画像処理し JSON データを生成する [tsukuba-gc](https://github.com/ysakasin/tsukuba-gc) が実装されています。  
しかし2024年現在では多少取り回しやすい Excel ファイルも公式から公開されているため、より簡易にデータを生成するスクリプトを新たに実装しました。

## Installation

### Dependencies
```
❯ python-V                            
Python 3.12.3
```
```
dependencies = [
    "openpyxl>=3.1.3",
    "ruff>=0.4.6",
]
```

### Project Setup & Run Script
- Clone
```
❯ git clone git@github.com:KyosukeHAGIWARA/tsukuba_garbage_calendar.git
❯ cd tsukuba_garbage_calendar
```

- Install Rye
```
❯ pip install rye
```

- Rye Sync
```
❯ rye sync
```

- Run Script
```
❯ . .venv/bin/activate
.venv ❯ cd scripts/
.venv ❯ sh process_calendar.sh
```

### Arguments
設定値は `scripts/process_calendar.sh` の中で切り替えられる他、 `scripts/generate_json_calendar_data.py` の実行時引数として渡すこともできる

|         設定値        |                 備考                |                       例                      |
|:---------------------:|:-----------------------------------:|:---------------------------------------------:|
|  EXCEL_FILE_DIRECTORY | Excelファイルが置かれたディレクトリ |          ../calendar_data/2024/xlsx/          |
|   EXCEL_FILE_FORMAT   |     Excelファイルのフォーマット     |                     *.xlsx                    |
|    EXCEL_SHEET_NAME   |       Excelファイルのシート名       |          ごみ出しパターン例外一括編集         |
|       START_DATE      |               開始日時              |                    20240401                   |
|        END_DATE       |               終了日時              |                    20250331                   |
| OUTPUT_JSON_FILE_NAME |        出力するJSONファイル名       | ../calendar_data/2024/json/calendar_data.json |