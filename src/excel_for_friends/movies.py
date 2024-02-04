SHEET_NAME = "Movies"
SHEET_INDEX = 0


class Movie:
    def __init__(self, information_color, success_color, end_color):
        self.information_color = information_color
        self.success_color = success_color
        self.end_color = end_color

    def _create_movie_sheet(self):
        self.wb.create_sheet(SHEET_NAME, SHEET_INDEX)

    def add_column_names(self):
        list_of_column_names = ["Name", "Genre", "Rating"]
        self._create_movie_sheet()
        self.wb[SHEET_NAME].append(list_of_column_names)

    def add_values_to_cells(self):
        sheet = self.wb[SHEET_NAME]
        next_row = sheet.max_row + 1
        movie_name = self._get_movie_name()
        movie_genre = self._get_movie_genre()
        sheet.cell(row=next_row, column=1).value = movie_name
        sheet.cell(row=next_row, column=2).value = movie_genre
        sheet.cell(row=next_row, column=3).value = (self._get_movie_rating() / 100)
        sheet.cell(row=next_row, column=3).number_format = '0%'
        if sheet.column_dimensions['A'].width > 11 or sheet.column_dimensions['B'].width > 11:
            sheet.column_dimensions['A'].width = max(len(movie_name), sheet.column_dimensions['A'].width)
            sheet.column_dimensions['B'].width = max(len(movie_genre), sheet.column_dimensions['B'].width)

    def _get_movie_name(self):
        movie_name = input(f"{self.success_color}Enter the name of the movie: {self.end_color}")
        return movie_name

    def _get_movie_genre(self):
        movie_genre = input(f"{self.success_color}Enter the genre of the movie: {self.end_color}")
        return movie_genre

    def _get_movie_rating(self):
        print(f"{self.information_color}The number have to be between 0 and 100. 100 is the best and 0 is the worst!{self.end_color}")
        movie_rating = int(input(f"{self.success_color}Enter your movie rating: {self.end_color}"))
        return movie_rating
