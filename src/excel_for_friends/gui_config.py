from src.excel_for_friends.excel_config import ExcelConfig
from src.excel_for_friends.exceptions import NumberNotInRange, EmptyFields
from tkinter import filedialog, ttk, messagebox
from openpyxl.utils.exceptions import InvalidFileException

import tkinter as tk
import os

LABEL_FONT = ("arial", 12, "bold")
RADIO_BUTTON_FONT = ("arial", 12, "bold")
CHECK_BOX_FONT = ("arial", 8, "bold")
ENTRY_FONT = ("arial", 8)
BUTTON_FONT = ("arial", 10, "bold")
BUTTON_COLOR = "#90ee90"
FRAME_BG_COLOR = "#f0f0f0"


class App(tk.Tk, ExcelConfig):
    def __init__(self):
        tk.Tk.__init__(self)
        self.file_path = None
        self.first_entry = None
        self.second_entry = None
        self.third_entry = None
        self.file_entry = None
        ExcelConfig.__init__(self, first_entry=self.first_entry, second_entry=self.second_entry, third_entry=self.third_entry, file_path=self.file_path)
        self.title("List of Hits")
        self.resizable(False, False)
        self.style = ttk.Style()
        self.style.configure("TButton", font=BUTTON_FONT, background=BUTTON_COLOR)
        self.first_label = None
        self.second_label = None
        self.third_label = None
        self.eText = tk.StringVar()
        self.radio_state = tk.IntVar(value=1)
        self.checked_state = tk.IntVar()
        self.create_widgets()
        self.protocol_close()

    def protocol_close(self):
        self.protocol("WM_DELETE_WINDOW", self.close_tkinter)

    def close_tkinter(self):
        if messagebox.askokcancel("Quit", "Do you want to quit?"):
            self._open_excel_file()
            self.save_excel_file()
            self.destroy()

    def create_widgets(self):
        self.first_row()
        self.second_row()
        self.third_row()
        self.fourth_row()
        self.fifth_row()
        self.sixth_row()

    def create_file(self):
        try:
            file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")], title="Create a file")
            self.file_path = file_path
            self._open_excel_file()
            self.wb.save(self.file_path)
            self.write_to_excel()
            end_name = os.path.basename(self.file_path)
            self.eText.set(end_name)
        except FileNotFoundError:
            messagebox.showerror("Error", "You have to create a file!")

    def load_file(self):
        try:
            file_path = filedialog.askopenfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")], title="Choose a file")
            self.file_path = file_path
            end_name = os.path.basename(self.file_path)
            self.eText.set(end_name)
            self._load_excel_file()
        except InvalidFileException:
            messagebox.showerror("Error", "You have to select a file!")

    def add_to_excel_button(self):
        try:
            self.choose_option()
            if self.get_checkbutton() == 0:
                self.quit()
            else:
                messagebox.showinfo("Status", "Progress was saved!")
            self.first_entry.focus()
            self.clear_entries()
        except ValueError:
            messagebox.showerror("Error", "The rating have to be in number format!")
        except KeyError:
            if len(str(self.file_entry.get()).strip()) == 0:
                messagebox.showerror("Error", "You have to select a file!")
            else:
                messagebox.showerror("Error", "You have to load another file! Or use the create button!")
        except PermissionError:
            messagebox.showerror("Error", "You have to close the Excel file first!")
        except NumberNotInRange:
            messagebox.showerror("Error", "Rating have to be in range 0 to 100!")
        except EmptyFields:
            messagebox.showerror("Error", "Fields should not be empty!")

    def change_labels(self):
        genre_list = ["Genre", "Genre", "Category", "Singer"]
        self.first_label['text'] = "Name:"
        if self.radio_state.get() == 1 or self.radio_state.get() == 2:
            self.second_label['text'] = f"{genre_list[self.radio_state.get() - 1]}:"
            self.second_label.pack(padx=12)
        elif self.radio_state.get() == 3:
            self.second_label['text'] = f"{genre_list[self.radio_state.get() - 1]}:"
            self.second_label.pack(padx=0)
        elif self.radio_state.get() == 4:
            self.second_label['text'] = f"{genre_list[self.radio_state.get() - 1]}:"
            self.second_label.pack(padx=10)
        self.third_label['text'] = "Rating:"

    def clear_entries(self):
        self.first_entry.delete(0, 'end')
        self.second_entry.delete(0, 'end')
        self.third_entry.delete(0, 'end')

    def get_checkbutton(self):
        return self.checked_state.get()

    def get_radio_button(self):
        return self.radio_state.get()

    def first_row(self):
        frame1 = tk.Frame(self, bg=FRAME_BG_COLOR)
        frame1.grid(row=0, column=0, padx=10, pady=10)
        load_button = ttk.Button(frame1, text="Open a File", command=self.load_file, takefocus=False)
        load_button.pack(side="left", padx=5, pady=5)
        create_button = ttk.Button(frame1, text="Create a File", command=self.create_file, takefocus=False)
        create_button.pack(side="left", padx=5, pady=5)
        self.file_entry = tk.Entry(frame1, font=LABEL_FONT, state="readonly", textvariable=self.eText, width=55)
        self.file_entry.pack(side="left", padx=5, pady=5)

    def second_row(self):
        frame2 = tk.Frame(self, bg=FRAME_BG_COLOR)
        frame2.grid(row=1, column=0, padx=10, pady=10)
        movies_radio_button = tk.Radiobutton(frame2, text="Movies", value=1, variable=self.radio_state, command=self.change_labels, font=RADIO_BUTTON_FONT, indicatoron=False)
        movies_radio_button.pack(side="left", padx=40, pady=5)
        shows_radio_button = tk.Radiobutton(frame2, text="TV Shows", value=2, variable=self.radio_state, command=self.change_labels, font=RADIO_BUTTON_FONT, indicatoron=False)
        shows_radio_button.pack(side="left", padx=40, pady=5)
        games_radio_button = tk.Radiobutton(frame2, text="Games", value=3, variable=self.radio_state, command=self.change_labels, font=RADIO_BUTTON_FONT, indicatoron=False)
        games_radio_button.pack(side="left", padx=40, pady=5)
        songs_radio_button = tk.Radiobutton(frame2, text="Songs", value=4, variable=self.radio_state, command=self.change_labels, font=RADIO_BUTTON_FONT, indicatoron=False)
        songs_radio_button.pack(side="left", padx=40, pady=5)

    def third_row(self):
        frame3 = tk.Frame(self, bg=FRAME_BG_COLOR)
        frame3.grid(row=2, column=0, padx=10, pady=10)
        self.first_label = tk.Label(frame3, text="Name:", font=LABEL_FONT)
        self.first_label.pack(side="left", padx=13)
        self.first_entry = tk.Entry(frame3, width=100)
        self.first_entry.pack(side="left", padx=5, pady=5)

    def fourth_row(self):
        frame4 = tk.Frame(self, bg=FRAME_BG_COLOR)
        frame4.grid(row=3, column=0, padx=10, pady=10)
        self.second_label = tk.Label(frame4, text="Genre:", font=LABEL_FONT)
        self.second_label.pack(side="left", padx=12)
        self.second_entry = tk.Entry(frame4, width=100)
        self.second_entry.pack(side="left", padx=5, pady=5)

    def fifth_row(self):
        frame5 = tk.Frame(self, bg=FRAME_BG_COLOR)
        frame5.grid(row=4, column=0, padx=10, pady=10)
        self.third_label = tk.Label(frame5, text="Rating:", font=LABEL_FONT)
        self.third_label.pack(side="left", padx=10)
        self.third_entry = tk.Entry(frame5, width=100)
        self.third_entry.pack(side="left", padx=5, pady=5)

    def sixth_row(self):
        frame6 = tk.Frame(self, bg=FRAME_BG_COLOR)
        frame6.grid(row=5, column=0, padx=10, pady=10)
        add_button = ttk.Button(frame6, text="Add to the excel", command=self.add_to_excel_button, takefocus=False)
        add_button.pack(side="left", padx=5, pady=5)
        checkbutton = tk.Checkbutton(frame6, text="add another row?", variable=self.checked_state, font=CHECK_BOX_FONT)
        checkbutton.pack(side="left", padx=5, pady=5)
