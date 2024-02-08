import pandas as pd
from openpyxl.styles import Font, Alignment, NamedStyle


class SortExcel:
    def __init__(self):
        self.list_with_sheets = ["Movies", "TV Shows", "Games", "Songs"]

    def sort_value(self, option_number):

        if self.list_with_sheets[option_number - 1] in self.wb.sheetnames:
            df = pd.read_excel(self.file_name, sheet_name=self.list_with_sheets[option_number - 1])

            df_sorted = df.sort_values(by='Rating', ascending=False)

            with pd.ExcelWriter(self.file_name, engine='openpyxl', mode='a', if_sheet_exists="overlay") as writer:
                df_sorted.to_excel(writer, sheet_name=self.list_with_sheets[option_number - 1], index=False)
                ws = writer.sheets[self.list_with_sheets[option_number - 1]]
                for row in ws.iter_rows(min_row=1, max_row=1):
                    for cell in row:
                        cell.font = Font(bold=False)
                        cell.alignment = Alignment(horizontal='left')
                        cell.border = None
