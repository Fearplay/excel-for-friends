from openpyxl.styles import PatternFill

from movies import Movie
from tv_shows import Show
from games import Game
from songs import Song

import openpyxl
import os

RED_FILL_COLOR = PatternFill(start_color='FFFF0000',
                             end_color='FFFF0000',
                             fill_type='solid')


class ExcelConfig(Movie, Show, Game, Song):
    def __init__(self, file_path, warning_color, information_color, success_color, end_color):
        Movie.__init__(self, information_color, success_color, end_color, RED_FILL_COLOR)
        Show.__init__(self, information_color, success_color, end_color, RED_FILL_COLOR)
        Game.__init__(self, information_color, success_color, end_color, RED_FILL_COLOR)
        Song.__init__(self, information_color, success_color, end_color, RED_FILL_COLOR)
        self.file_path = file_path
        self.file_name = "list_of_hits.xlsx"
        self.warning_color = warning_color
        self.information_color = information_color
        self.success_color = success_color
        self.end_color = end_color
        self.wb = openpyxl.Workbook()

    def _open_excel_file(self):
        self.wb = self.wb

    def _load_excel_file(self):
        self.wb = openpyxl.load_workbook(self.file_name)

    def write_to_excel(self):
        if os.path.exists(self.file_path + f"{self.file_name}"):
            self._load_excel_file()
        else:
            self._open_excel_file()
            Movie.add_column_names(self)
            Show.add_column_names(self)
            Game.add_column_names(self)
            Song.add_column_names(self)

        self.choose_option()
        self.safe_excel_file()

    def choose_option(self):
        answer = "y"
        option = self._get_option_number()
        while int != type(option) or option < 1 or option > 5:
            if int == type(option):
                print(f"{self.warning_color}The number have to be between 1 and 4!{self.end_color}")
            option = self._get_option_number()

        while answer == "y" or answer == "yes":
            if option == 1:
                Movie.add_values_to_cells(self)
            if option == 2:
                Show.add_values_to_cells(self)
                self.safe_excel_file()
            if option == 3:
                Game.add_values_to_cells(self)
            if option == 4:
                Song.add_values_to_cells(self)
            answer = self._get_answer()

    def _get_option_number(self):
        try:
            option = int(input(f"{self.success_color}Choose what do you want to add 1 - movie, 2 - tv show, 3 - game, 4 - song: {self.end_color}"))
            return option
        except ValueError:
            print(f"{self.warning_color}You have to type a number!{self.end_color}")

    def _get_answer(self):
        try:
            answer = input(f"{self.success_color}Do you want to add another row?(y/n): {self.end_color}")
            return answer.lower()
        except ValueError:
            print(f"{self.warning_color}You have to type a number!{self.end_color}")

    def delete_sheet(self):
        if 'Sheet' in self.wb.sheetnames:
            self.wb.remove(self.wb['Sheet'])

    def safe_excel_file(self):
        self.delete_sheet()
        self.wb.save(self.file_name)
        print(f"{self.success_color}Changes were saved!{self.end_color}")
