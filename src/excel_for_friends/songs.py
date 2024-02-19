from src.excel_for_friends.exceptions import NumberNotInRange, EmptyFields

SHEET_NAME = "Songs"
SHEET_INDEX = 3


class Song:
    def __init__(self,first_entry, second_entry, third_entry, fill_color):
        self.first_entry = first_entry
        self.second_entry = second_entry
        self.third_entry = third_entry
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
        song_name = str(self.first_entry.get())
        if len(song_name.strip()) == 0:
            raise EmptyFields
        else:
            return song_name

    def _get_singer_name(self):
        song_genre = str(self.second_entry.get())
        if len(song_genre.strip()) == 0:
            raise EmptyFields
        else:
            return song_genre

    def _get_song_rating(self):
        song_rating = int(self.third_entry.get())
        if song_rating < 0 or song_rating > 100:
            raise NumberNotInRange
        else:
            return song_rating
