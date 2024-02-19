from src.excel_for_friends.exceptions import NumberNotInRange, EmptyFields

SHEET_NAME = "Games"
SHEET_INDEX = 2


class Game:
    def __init__(self, first_entry, second_entry, third_entry, fill_color):
        self.first_entry = first_entry
        self.second_entry = second_entry
        self.third_entry = third_entry
        self.fill_color = fill_color

    def _create_game_sheet(self):
        self.wb.create_sheet(SHEET_NAME, SHEET_INDEX)

    def add_column_names(self):
        list_of_column_names = ["Name", "Category", "Rating", "Played"]
        self._create_game_sheet()
        self.wb[SHEET_NAME].append(list_of_column_names)

    def add_values_to_cells(self):
        sheet = self.wb[SHEET_NAME]
        next_row = sheet.max_row + 1
        game_name = self._get_game_name()
        game_genre = self._get_game_genre()
        sheet.cell(row=next_row, column=1).value = game_name
        sheet.cell(row=next_row, column=2).value = game_genre
        sheet.cell(row=next_row, column=3).value = (self._get_game_rating() / 100)
        sheet.cell(row=next_row, column=3).number_format = '0%'
        sheet.cell(row=next_row, column=4).fill = self.fill_color
        if sheet.column_dimensions['A'].width > 11 or sheet.column_dimensions['B'].width > 11:
            sheet.column_dimensions['A'].width = max(len(game_name), sheet.column_dimensions['A'].width)
            sheet.column_dimensions['B'].width = max(len(game_genre), sheet.column_dimensions['B'].width)

    def _get_game_name(self):
        game_name = str(self.first_entry.get())
        if len(game_name.strip()) == 0:
            raise EmptyFields
        else:
            return game_name

    def _get_game_genre(self):
        game_genre = str(self.second_entry.get())
        if len(game_genre.strip()) == 0:
            raise EmptyFields
        else:
            return game_genre

    def _get_game_rating(self):
        game_rating = int(self.third_entry.get())
        if game_rating < 0 or game_rating > 100:
            raise NumberNotInRange
        else:
            return game_rating
