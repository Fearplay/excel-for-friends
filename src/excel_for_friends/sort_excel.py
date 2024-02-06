import pandas as pd


class SortExcel:
    def __init__(self):
        pass

    def sort_value(self, sheet_name):
        if sheet_name in self.wb.sheetnames:
            df = pd.read_excel(self.file_name, sheet_name=sheet_name)

            df_sorted = df.sort_values(by='Rating', ascending=False)

            with pd.ExcelWriter(self.file_name, engine='openpyxl', mode='a', if_sheet_exists="overlay") as writer:
                df_sorted.to_excel(writer, sheet_name=sheet_name, index=False)
