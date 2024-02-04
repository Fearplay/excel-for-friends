SHEET_NAME = "TV Shows"
SHEET_INDEX = 1


class Show:
    def __init__(self, information_color, success_color, end_color):
        self.information_color = information_color
        self.success_color = success_color
        self.end_color = end_color

    def _create_show_sheet(self):
        self.wb.create_sheet(SHEET_NAME, SHEET_INDEX)

    def add_column_names(self):
        list_of_column_names = ["Name", "Genre", "Rating"]
        self._create_show_sheet()
        self.wb[SHEET_NAME].append(list_of_column_names)

    def add_values_to_cells(self):
        sheet = self.wb[SHEET_NAME]
        next_row = sheet.max_row + 1
        show_name = self._get_show_name()
        show_genre = self._get_show_genre()
        sheet.cell(row=next_row, column=1).value = show_name
        sheet.cell(row=next_row, column=2).value = show_genre
        sheet.cell(row=next_row, column=3).value = (self._get_show_rating() / 100)
        sheet.cell(row=next_row, column=3).number_format = '0%'
        if sheet.column_dimensions['A'].width > 11 or sheet.column_dimensions['B'].width > 11:
            sheet.column_dimensions['A'].width = max(len(show_name), sheet.column_dimensions['A'].width)
            sheet.column_dimensions['B'].width = max(len(show_genre), sheet.column_dimensions['B'].width)

    def _get_show_name(self):
        show_name = input(f"{self.success_color}Enter the name of the TV Show: {self.end_color}")
        return show_name

    def _get_show_genre(self):
        show_genre = input(f"{self.success_color}Enter the genre of the TV Show: {self.end_color}")
        return show_genre

    def _get_show_rating(self):
        print(f"{self.information_color}The number have to be between 0 and 100. 100 is the best and 0 is the worst!{self.end_color}")
        show_rating = int(input(f"{self.success_color}Enter your TV Show rating: {self.end_color}"))
        return show_rating
