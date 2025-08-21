import xlwings as xw
from tkinter import Tk, Button, Label
from tkcalendar import Calendar

def show_calendar(cell_address):

    # Validate cell_address
    if not cell_address or not isinstance(cell_address, str):
        print("Error: Invalid cell address provided")
        return

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
        wb = xw.Book.caller()
        ws = wb.sheets.active
        ws.api.Unprotect()

        selected_date = cal.get_date()
        result_label.config(text=selected_date)

        ws.range(cell_address).value = selected_date

        # Reprotect the sheet if it was protected
        if ws.api.ProtectContents is False:
            ws.api.Protect()

        root.destroy()

    btn = Button(root, text="Confirm Date", command=grab_date)
    btn.pack(pady=10)

    root.mainloop()
