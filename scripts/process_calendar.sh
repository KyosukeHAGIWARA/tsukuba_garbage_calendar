# Excelファイルが置かれたディレクトリ: ex. ../calendar_data/2024/xlsx/
EXCEL_FILE_DIRECTORY="../calendar_data/2024/xlsx/"
# Excelファイルのフォーマット: ex. *.xlsx
EXCEL_FILE_FORMAT="*.xlsx"
# Excelファイルのシート名: ex. ごみ出しパターン例外一括編集
EXCEL_SHEET_NAME="ごみ出しパターン例外一括編集"
# カレンダー開始日時: ex. 20240401
START_DATE="20240401"
# カレンダー終了日時: ex. 20250331
END_DATE="20250331"
# 出力するJSONファイル名: ex. ../calendar_data/2024/json/calendar_data.json
OUTPUT_JSON_FILE_NAME="../calendar_data/2024/json/calendar_data.json"

python generate_json_calendar_data.py \
    --excel_dir $EXCEL_FILE_DIRECTORY \
    --excel_name_format $EXCEL_FILE_FORMAT \
    --sheet_name $EXCEL_SHEET_NAME \
    --start $START_DATE \
    --end $END_DATE \
    --output_json_file_name $OUTPUT_JSON_FILE_NAME
