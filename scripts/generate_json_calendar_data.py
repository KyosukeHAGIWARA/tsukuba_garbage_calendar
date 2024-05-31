import openpyxl
import os
import glob
from enum import Enum
import logging
import json
from datetime import datetime, timedelta

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

EXCEL_FILE_DIRECTORY = "../calendar_data/2024/xlsx/"
EXCEL_FILE_FORMAT = "*.xlsx"
EXCEL_SHEET_NAME = "ごみ出しパターン例外一括編集"

START_DATE = datetime(2024, 4, 1)
END_DATE = datetime(2025, 3, 31)

SUBJECT_BLOCK_LIST_KEY = "subject_block_list"
SUBJECT_BLOCK_KEY = "subject_block"
SUBJECT_BLOCK_PRONUNCIATION_KEY = "subject_block_pronunciation"
CALENDAR_KEY = "calendar"

TRUE_KEY = "true"
FALSE_KEY = "false"

OUTPUT_JSON_FILE_NAME = "../calendar_data/2024/json/calendar_data.json"


class GARBAGE_TYPE(Enum):
    BURNABLE = "燃やせるごみ"
    BOTTLE = "びん"
    SPRAY = "スプレー容器"
    PET = "ペットボトル"
    NON_BURNABLE = "燃やせないごみ"
    PAPER_CLOTH = "古紙・古布"
    PLASTIC = "プラスチック製容器包装"
    CAN = "かん"
    BULKY_WASTE = "粗大ごみ（予約制）"

    @classmethod
    def show_all(cls) -> list[str]:
        return list(map(lambda c: c.value, cls))


def app():
    # シートのデータを辞書としてロード
    calendar_data = {}
    """calendar_data は以下の形式でデータを保持する
    {
        '地区エリアA':{
            'subject_block_list':{
                '赤塚': {
                    'subject_block': '赤塚',
                    'subject_block_pronunciation': 'あかつか'
                },
                '青塚': {...}
            },
            'calendar':{
                '2024/04/01':{
                    GARBAGE_TYPE.BURNABLE: 'true',
                    GARBAGE_TYPE.BOTTLE: 'true',
                    GARBAGE_TYPE.SPRAY: 'true',
                    GARBAGE_TYPE.PET: 'true',
                    GARBAGE_TYPE.NON_BURNABLE: 'true',
                    GARBAGE_TYPE.PAPER_CLOTH: 'true',
                    GARBAGE_TYPE.PLASTIC: 'true',
                    GARBAGE_TYPE.CAN: 'true',
                    GARBAGE_TYPE.BULKY_WASTE: 'false',
                },
                {...}
            }
        },
        '地区エリアB':{...}
    }
    """

    # 指定したディレクトリ内のすべてのxlsxファイルを開く
    for monthly_file_name in glob.glob(
        os.path.join(EXCEL_FILE_DIRECTORY, EXCEL_FILE_FORMAT)
    ):
        logger.info(f"processing excel file: {monthly_file_name}")
        # ワークブックを開く
        workbook = openpyxl.load_workbook(monthly_file_name)
        # シートを開く
        sheet = workbook[EXCEL_SHEET_NAME]
        rows = sheet.iter_rows(values_only=True)

        # シートのデータを読み取って calendar_data に格納
        analyze_sheet_data(rows, calendar_data)
    else:
        logger.info("finish processing all excel files")
        filled_calendar_data = fill_in_no_pickup_date(calendar_data)

        logger.info(f"output to json file: {OUTPUT_JSON_FILE_NAME}")
        output_json_file(OUTPUT_JSON_FILE_NAME, filled_calendar_data)
        logger.info("output finished")


def analyze_sheet_data(rows, calendar_data):
    # 1行目の header 行からゴミの種類の並び順を取得
    first_row = list(next(rows))
    garbage_header_list = generate_garbage_header(first_row)
    logger.debug(f"garbage_header_list: {garbage_header_list}")

    for row in rows:
        subject_block = row[0]

        address_annotation = row[1].split()
        logger.debug(address_annotation)
        subject_area = address_annotation[0]
        subject_block_pronunciation = address_annotation[1]
        logger.debug("")
        logger.debug(
            f"subject_area: {subject_area}, subject_block: {subject_block}, subject_block_pronunciation: {subject_block_pronunciation}"
        )

        # insert subject_block data
        if subject_area not in calendar_data:
            calendar_data[subject_area] = {
                SUBJECT_BLOCK_LIST_KEY: {},
                CALENDAR_KEY: {},
            }
        if subject_block not in calendar_data.get(subject_area, {}).get(
            SUBJECT_BLOCK_LIST_KEY, {}
        ):
            calendar_data[subject_area][SUBJECT_BLOCK_LIST_KEY][subject_block] = {
                SUBJECT_BLOCK_KEY: subject_block,
                SUBJECT_BLOCK_PRONUNCIATION_KEY: subject_block_pronunciation,
            }

        date_list = list(row[2:])
        for index, date_item in enumerate(date_list):
            current_garbage_type = garbage_header_list[index]
            date_list = date_item.split(",")
            logger.debug(
                f"garbage_type: {current_garbage_type}, date_item: {date_list}"
            )

            # insert calendar data
            for date in date_list:
                if date not in calendar_data[subject_area][CALENDAR_KEY]:
                    calendar_data[subject_area][CALENDAR_KEY][date] = (
                        blank_calendar_item()
                    )

                calendar_data[subject_area][CALENDAR_KEY][date][
                    current_garbage_type.value
                ] = TRUE_KEY


def generate_garbage_header(header_list):
    garbage_header_list = []
    for i, header_item in enumerate(header_list[2:]):
        if header_item in GARBAGE_TYPE.show_all():
            garbage_header_list.append(GARBAGE_TYPE(header_item))
        else:
            logger.error(f"Error: index: {i}, {header_item} is not in GARBAGE_TYPE")
            garbage_header_list.append(None)
    return garbage_header_list


def blank_calendar_item():
    calendar_blank = {}
    for garbage_type in GARBAGE_TYPE:
        calendar_blank[garbage_type.value] = FALSE_KEY

    return calendar_blank


def fill_in_no_pickup_date(calendar_data):
    for subject_area in calendar_data:
        current_date = START_DATE
        while current_date <= END_DATE:
            current_date_str = current_date.strftime("%Y/%m/%d")
            if current_date_str not in calendar_data[subject_area][CALENDAR_KEY]:
                calendar_data[subject_area][CALENDAR_KEY][current_date_str] = (
                    blank_calendar_item()
                )
                logger.debug(f"fill in no pickup date: {current_date_str}")
            current_date += timedelta(days=1)

    return calendar_data


def output_json_file(file_name, data):
    # JSONファイルに出力
    with open(file_name, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4, sort_keys=True)


if __name__ == "__main__":
    app()
