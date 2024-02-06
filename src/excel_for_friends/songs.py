SHEET_NAME = "Songs"
SHEET_INDEX = 3


class Song:
    def __init__(self, information_color, success_color, end_color, fill_color):
        self.information_color = information_color
        self.success_color = success_color
        self.end_color = end_color
        self.fill_color = fill_color

    def _create_song_sheet(self):
        self.wb.create_sheet(SHEET_NAME, SHEET_INDEX)

    def add_column_names(self):
        list_of_column_names = ["Name", "Singer", "Rating", "Heard"]
        self._create_song_sheet()
        self.wb[SHEET_NAME].append(list_of_column_names)

    def add_values_to_cells(self):
        sheet = self.wb[SHEET_NAME]
        next_row = sheet.max_row + 1
        song_name = self._get_song_name()
        singer_name = self._get_singer_name()
        sheet.cell(row=next_row, column=1).value = song_name
        sheet.cell(row=next_row, column=2).value = singer_name
        sheet.cell(row=next_row, column=3).value = (self._get_song_rating() / 100)
        sheet.cell(row=next_row, column=3).number_format = '0%'
        sheet.cell(row=next_row, column=4).fill = self.fill_color
        if sheet.column_dimensions['A'].width > 11 or sheet.column_dimensions['B'].width > 11:
            sheet.column_dimensions['A'].width = max(len(song_name), sheet.column_dimensions['A'].width)
            sheet.column_dimensions['B'].width = max(len(singer_name), sheet.column_dimensions['B'].width)

    def _get_song_name(self):
        song_name = input(f"{self.success_color}Enter the name of the song: {self.end_color}")
        return song_name

    def _get_singer_name(self):
        song_genre = input(f"{self.success_color}Enter the name of the singer: {self.end_color}")
        return song_genre

    def _get_song_rating(self):
        print(f"{self.information_color}The number have to be between 0 and 100. 100 is the best and 0 is the worst!{self.end_color}")
        song_rating = int(input(f"{self.success_color}Enter your song rating: {self.end_color}"))
        return song_rating
