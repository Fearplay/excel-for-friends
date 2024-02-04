SHEET_NAME = "Games"
SHEET_INDEX = 2


class Game:
    def __init__(self, information_color, success_color, end_color):
        self.information_color = information_color
        self.success_color = success_color
        self.end_color = end_color

    def _create_game_sheet(self):
        self.wb.create_sheet(SHEET_NAME, SHEET_INDEX)

    def add_column_names(self):
        list_of_column_names = ["Name", "Category", "Rating"]
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
        if sheet.column_dimensions['A'].width > 11 or sheet.column_dimensions['B'].width > 11:
            sheet.column_dimensions['A'].width = max(len(game_name), sheet.column_dimensions['A'].width)
            sheet.column_dimensions['B'].width = max(len(game_genre), sheet.column_dimensions['B'].width)

    def _get_game_name(self):
        game_name = input(f"{self.success_color}Enter the name of the game: {self.end_color}")
        return game_name

    def _get_game_genre(self):
        game_genre = input(f"{self.success_color}Enter the genre of the game: {self.end_color}")
        return game_genre

    def _get_game_rating(self):
        print(f"{self.information_color}The number have to be between 0 and 100. 100 is the best and 0 is the worst!{self.end_color}")
        game_rating = int(input(f"{self.success_color}Enter your game rating: {self.end_color}"))
        return game_rating
