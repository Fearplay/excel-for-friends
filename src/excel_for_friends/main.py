from src.excel_for_friends.gui_config import App

if __name__ == "__main__":
    app = App()
    app.eval('tk::PlaceWindow . center')
    app.mainloop()
