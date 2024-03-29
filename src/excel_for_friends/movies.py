from src.excel_for_friends.exceptions import NumberNotInRange, EmptyFields

SHEET_NAME = "Movies"
SHEET_INDEX = 0


class Movie:
    def __init__(self, first_entry, second_entry, third_entry, fill_color):
        self.first_entry = first_entry
        self.second_entry = second_entry
        self.third_entry = third_entry
        self.fill_color = fill_color

    def _create_movie_sheet(self):
        self.wb.create_sheet(SHEET_NAME, SHEET_INDEX)

    def add_column_names(self):
        list_of_column_names = ["Name", "Genre", "Rating", "Watched"]
        self._create_movie_sheet()
        self.wb[SHEET_NAME].append(list_of_column_names)

    def add_values_to_cells(self):
        sheet = self.wb[SHEET_NAME]
        next_row = sheet.max_row + 1
        movie_name = self._get_movie_name()
        movie_genre = self._get_movie_genre()
        movie_rating = self._get_movie_rating()
        if movie_rating < 0 or movie_rating > 100:
            raise NumberNotInRange
        else:
            sheet.cell(row=next_row, column=1).value = movie_name
            sheet.cell(row=next_row, column=2).value = movie_genre
            sheet.cell(row=next_row, column=3).value = (self._get_movie_rating() / 100)
            sheet.cell(row=next_row, column=3).number_format = '0%'
            sheet.cell(row=next_row, column=4).fill = self.fill_color
            if sheet.column_dimensions['A'].width > 11 or sheet.column_dimensions['B'].width > 11:
                sheet.column_dimensions['A'].width = max(len(movie_name), sheet.column_dimensions['A'].width)
                sheet.column_dimensions['B'].width = max(len(movie_genre), sheet.column_dimensions['B'].width)

    def _get_movie_name(self):
        movie_name = str(self.first_entry.get())
        if len(movie_name.strip()) == 0:
            raise EmptyFields
        else:
            return movie_name

    def _get_movie_genre(self):
        movie_genre = str(self.second_entry.get())
        if len(movie_genre.strip()) == 0:
            raise EmptyFields
        else:
            return movie_genre

    def _get_movie_rating(self):
        movie_rating = int(self.third_entry.get())
        return movie_rating
