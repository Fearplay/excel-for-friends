import pandas as pd
from openpyxl.styles import Font, Alignment


class SortExcel:
    def __init__(self):
        self.list_with_sheets = ["Movies", "TV Shows", "Games", "Songs"]

    def sort_value(self):
        for sheet_name in self.list_with_sheets:
            if sheet_name in self.wb.sheetnames:
                df = pd.read_excel(self.file_path, sheet_name=sheet_name)
                df_sorted = df.sort_values(by=['Rating', 'Name'], ascending=[False, True])
                with pd.ExcelWriter(self.file_path, engine='openpyxl', mode='a', if_sheet_exists="overlay") as writer:
                    df_sorted.to_excel(writer, sheet_name=sheet_name, index=False)
                    ws = writer.sheets[sheet_name]
                    for row in ws.iter_rows(min_row=1, max_row=1):
                        for cell in row:
                            cell.font = Font(bold=False)
                            cell.alignment = Alignment(horizontal='left')
                            cell.border = None
