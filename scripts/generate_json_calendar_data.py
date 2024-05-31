import argparse
import openpyxl
import os
import glob
from enum import Enum
import logging
import json
from datetime import datetime, timedelta

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class GarbageType(Enum):
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


class GarbageCalendarProcesser:
    def __init__(
        self,
        excel_file_directory: str,
        excel_file_format: str,
        excel_sheet_name: str,
        start_date: str,
        end_date: str,
        output_json_file_name: str,
        subject_block_list_key="subject_block_list",
        subject_block_key="subject_block",
        subject_block_pronunciation_key="subject_block_pronunciation",
        calendar_key="calendar",
        true_key="true",
        false_key="false",
    ):
        self.excel_file_directory = excel_file_directory
        self.excel_file_format = excel_file_format
        self.excel_sheet_name = excel_sheet_name
        self.start_date = start_date
        self.end_date = end_date
        self.output_json_file_name = output_json_file_name

        self.calendar_data = {}
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

        # you can change the key name for output json if you want
        self.subject_block_list_key = subject_block_list_key
        self.subject_block_key = subject_block_key
        self.subject_block_pronunciation_key = subject_block_pronunciation_key
        self.calendar_key = calendar_key
        self.true_key = true_key
        self.false_key = false_key

    def process_calender(self):
        # シートのデータを辞書としてロード

        # 指定したディレクトリ内のすべてのxlsxファイルを開く
        for monthly_file_name in glob.glob(
            os.path.join(self.excel_file_directory, self.excel_file_format)
        ):
            logger.info(f"processing excel file: {monthly_file_name}")
            # ワークブックを開く
            workbook = openpyxl.load_workbook(monthly_file_name)
            # シートを開く
            sheet = workbook[self.excel_sheet_name]
            rows = sheet.iter_rows(values_only=True)

            # シートのデータを読み取って calendar_data に格納
            self.__analyze_sheet_data(rows, self.calendar_data)
        else:
            logger.info("finish processing all excel files")
            filled_calendar_data = self.__fill_in_no_pickup_date(self.calendar_data)

            logger.info(f"output to json file: {self.output_json_file_name}")
            self.__output_json_file(self.output_json_file_name, filled_calendar_data)
            logger.info("output finished")

    def __analyze_sheet_data(self, rows, calendar_data):
        # 1行目の header 行からゴミの種類の並び順を取得
        first_row = list(next(rows))
        garbage_header_list = self.__generate_garbage_header(first_row)
        logger.debug(f"garbage_header_list: {garbage_header_list}")

        # 2行目以降のデータを読み取る
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
                    self.subject_block_list_key: {},
                    self.calendar_key: {},
                }
            if subject_block not in calendar_data.get(subject_area, {}).get(
                self.subject_block_list_key, {}
            ):
                calendar_data[subject_area][self.subject_block_list_key][
                    subject_block
                ] = {
                    self.subject_block_key: subject_block,
                    self.subject_block_pronunciation_key: subject_block_pronunciation,
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
                    if date not in calendar_data[subject_area][self.calendar_key]:
                        calendar_data[subject_area][self.calendar_key][date] = (
                            self.__blank_calendar_item()
                        )

                    calendar_data[subject_area][self.calendar_key][date][
                        current_garbage_type.value
                    ] = self.true_key

    def __generate_garbage_header(self, header_list):
        garbage_header_list = []
        for i, header_item in enumerate(header_list[2:]):
            if header_item in GarbageType.show_all():
                garbage_header_list.append(GarbageType(header_item))
            else:
                logger.error(f"Error: index: {i}, {header_item} is not in GarbageType")
                garbage_header_list.append(None)
        return garbage_header_list

    def __blank_calendar_item(self):
        calendar_blank = {}
        for garbage_type in GarbageType:
            calendar_blank[garbage_type.value] = self.false_key

        return calendar_blank

    def __fill_in_no_pickup_date(self, calendar_data):
        for subject_area in calendar_data:
            current_date = self.start_date
            while current_date <= self.end_date:
                current_date_str = current_date.strftime("%Y/%m/%d")
                if (
                    current_date_str
                    not in calendar_data[subject_area][self.calendar_key]
                ):
                    calendar_data[subject_area][self.calendar_key][current_date_str] = (
                        self.__blank_calendar_item()
                    )
                    logger.debug(f"fill in no pickup date: {current_date_str}")
                current_date += timedelta(days=1)

        return calendar_data

    def __output_json_file(self, file_name, data):
        # JSONファイルに出力
        with open(file_name, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=4, sort_keys=True)


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--excel_dir",
        type=str,
        required=True,
        help="Excelファイルが置かれたディレクトリ: ex. ../calendar_data/2024/xlsx/",
    )
    parser.add_argument(
        "--excel_name_format",
        type=str,
        required=True,
        help="Excelファイル名のフォーマット: ex. *.xlsx",
    )
    parser.add_argument(
        "--sheet_name",
        type=str,
        required=True,
        help="Excelファイルのシート名: ex. ごみ出しパターン例外一括編集",
    )
    parser.add_argument(
        "--start",
        type=str,
        required=True,
        help="カレンダー開始日時: ex. 20240401",
    )
    parser.add_argument(
        "--end",
        type=str,
        required=True,
        help="カレンダー終了日時: ex. 20250331",
    )
    parser.add_argument(
        "--output_json_file_name",
        type=str,
        required=True,
        help="出力するJSONファイル名: ex. ../calendar_data/2024/json/calendar_data.json",
    )

    params = parser.parse_args()
    try:
        start_date = datetime.strptime(params.start, "%Y%m%d")
        end_date = datetime.strptime(params.end, "%Y%m%d")
    except ValueError:
        logger.error("Error: start or end date is invalid format")
        logger.error(f"current start date: {params.start}")
        logger.error(f"current end date: {params.end}")
        exit(1)

    app = GarbageCalendarProcesser(
        params.excel_dir,
        params.excel_name_format,
        params.sheet_name,
        start_date,
        end_date,
        params.output_json_file_name,
    )
    app.process_calender()
