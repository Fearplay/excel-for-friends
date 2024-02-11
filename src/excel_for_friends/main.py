from src.excel_for_friends.excel_config import ExcelConfig

YELLOW_START = '\033[33m'
GREEN_START = '\033[92m'
RED_START = '\033[91m'
COLOR_END = '\033[00m'

excel_work = ExcelConfig(file_path="../../data/output/", file_name="list_of_hits.xlsx", warning_color=RED_START, information_color=YELLOW_START, success_color=GREEN_START,
                         end_color=COLOR_END)

if __name__ == "__main__":
    excel_work.write_to_excel()
