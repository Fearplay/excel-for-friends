from src.excel_for_friends.exceptions import NumberNotInRange, EmptyFields

SHEET_NAME = "TV Shows"
SHEET_INDEX = 1


class Show:
    def __init__(self, first_entry, second_entry, third_entry, fill_color):
        self.first_entry = first_entry
        self.second_entry = second_entry
        self.third_entry = third_entry
        self.fill_color = fill_color

    def _create_show_sheet(self):
        self.wb.create_sheet(SHEET_NAME, SHEET_INDEX)

    def add_column_names(self):
        list_of_column_names = ["Name", "Genre", "Rating", "Watched"]
        self._create_show_sheet()
        self.wb[SHEET_NAME].append(list_of_column_names)

    def add_values_to_cells(self):
        sheet = self.wb[SHEET_NAME]
        next_row = sheet.max_row + 1
        show_name = self._get_show_name()
        show_genre = self._get_show_genre()
        show_rating = self._get_show_rating()
        if show_rating < 0 or show_rating > 100:
            raise NumberNotInRange
        else:
            sheet.cell(row=next_row, column=1).value = show_name
            sheet.cell(row=next_row, column=2).value = show_genre
            sheet.cell(row=next_row, column=3).value = (self._get_show_rating() / 100)
            sheet.cell(row=next_row, column=3).number_format = '0%'
            sheet.cell(row=next_row, column=4).fill = self.fill_color
            if sheet.column_dimensions['A'].width > 11 or sheet.column_dimensions['B'].width > 11:
                sheet.column_dimensions['A'].width = max(len(show_name), sheet.column_dimensions['A'].width)
                sheet.column_dimensions['B'].width = max(len(show_genre), sheet.column_dimensions['B'].width)

    def _get_show_name(self):
        show_name = str(self.first_entry.get())
        if len(show_name.strip()) == 0:
            raise EmptyFields
        else:
            return show_name

    def _get_show_genre(self):
        show_genre = str(self.second_entry.get())
        if len(show_genre.strip()) == 0:
            raise EmptyFields
        else:
            return show_genre

    def _get_show_rating(self):
        show_rating = int(self.third_entry.get())
        return show_rating
