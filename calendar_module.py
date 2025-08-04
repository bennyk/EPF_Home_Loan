import xlwings as xw
from tkinter import Tk, Button, Label
from tkcalendar import Calendar


def show_calendar(cell_address):
    root = Tk()
    x = root.winfo_pointerx()
    y = root.winfo_pointery()
    root.geometry(f"+{x}+{y}")

    root.title("Select Date")

    cal = Calendar(root, selectmode='day')
    cal.pack(pady=10)

    result_label = Label(root, text="")
    result_label.pack(pady=10)

    def grab_date():
        selected_date = cal.get_date()
        result_label.config(text=selected_date)

        wb = xw.Book.caller()
        ws = wb.sheets.active
        ws.range(cell_address).value = selected_date

        root.destroy()

    btn = Button(root, text="Confirm Date", command=grab_date)
    btn.pack(pady=10)

    root.mainloop()
