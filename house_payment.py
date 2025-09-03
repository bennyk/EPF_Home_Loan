import xlwings as xw
from datetime import datetime, timedelta

verbose = 0

# Mapping text to column
my_dict = {"Type": 'A', "Invoice date": 'E', "Due date": 'F', "Next payment": 'G', "Days": 'H',
           "Date paid": 'I', "State flag": 'J', "Status [4]": 'K' }


def run():
    try:
        wb = xw.Book("house_payment.xlsm")
        ws = wb.sheets["Estimation Expenses"]
        ws.api.Unprotect()

        # TODO xlwings sheets does not implement protect or unprotect sheets. Hmm
        # ws.unprotect()
        # ws.protect(password="xlwings")

        print(f"Opened workbook: {wb.name}")
        for i in range(2, 33):
            if ws.range(f"{my_dict["Type"]}{i}").value is None:
                continue

            if ws.range(f"{my_dict["Type"]}{i}").value[0] == "🏠":
                print(f"ℹ️ {ws.range(f'{my_dict["Type"]}{i}').value}")
            else:
                check_update(wb, i)

        # wb.save(); wb.close()
        print("xlwings update completed.")

        # Reprotect the sheet if it was protected
        if ws.api.ProtectContents is False:
            ws.api.Protect()
    except Exception as e:
        print(f"Error in run: {e}")


def check_update(wb, row_num):
    # Update first before checking.
    update_next_invoice(wb, row_num)
    check_neutral(wb, row_num)


def check_neutral(wb, row_num):
    ws = wb.sheets["Estimation Expenses"]
    status = ws.range(f"{my_dict["Status [4]"]}{row_num}").value
    item_name = ws.range(f"{my_dict["Type"]}{row_num}").value
    # print(f"Checking row {row_num}, Status: {status}")

    # Target "Date Paid" cell (Column J)
    date_paid_cell = ws.range(f"{my_dict["Date paid"]}{row_num}")
    date_paid_cell.color = None

    if status == "Neutral":
        due_date = ws.range(f"{my_dict["Due date"]}{row_num}").value
        print(f"⚠️ {item_name}: Status is {status}. Due date is {due_date.strftime('%d-%m-%y')}. Consider payment.")

        # Clear any prefilled content and highlight the cell
        date_paid_cell.value = ""
        date_paid_cell.color = (255, 255, 153)  # Post-it yellow

    elif status == "Bad":
        print(f"📛 {item_name}: Status is {status}.")

        # Clear and highlight here too, if that's consistent with your system
        date_paid_cell.value = ""
        date_paid_cell.color = (255, 255, 153)


def update_next_invoice(wb, row_num):
    ws = wb.sheets["Estimation Expenses"]
    # print(f"Processing row {row_num}, Date Paid: {date_paid}")
    date_paid = ws.range(f"{my_dict["Date paid"]}{row_num}").value
    if date_paid is not None and date_paid != "":
        due_date = ws.range(f"{my_dict['Due date']}{row_num}").value
        days_value = ws.range(f"{my_dict['Days']}{row_num}").value
        if isinstance(days_value, (int, float)):
            # Check if due date is past (days_value < 0) or use direct comparison
            if days_value < 0 or (datetime.now() - due_date).days > 0:
                next_payment = ws.range(f"{my_dict['Next payment']}{row_num}").value
                if next_payment != "Nope":
                    assert isinstance(next_payment, datetime)
                    ws.range(f"{my_dict['Invoice date']}{row_num}").value = next_payment
                    ws.range(f"{my_dict['Invoice date']}{row_num}").number_format = "dd-mm-yy"
                    # # State flag and Next payment no longer needed with Excel formula
                    # ws.range(f"{my_dict['State flag']}{row_num}").value = ws.range(f"{my_dict['State flag']}{row_num}").value + 1
                    # ws.range(f"{my_dict['Next payment']}{row_num}").value = ""
                    ws.range(f"{my_dict['Date paid']}{row_num}").value = ""
                    formatted_date = (ws.range(f"{my_dict['Invoice date']}{row_num}").value.strftime("%d-%m-%y")
                                    if isinstance(ws.range(f"{my_dict['Invoice date']}{row_num}").value, datetime) else "N/A")
                    print(f"{ws.range(f'{my_dict["Type"]}{row_num}').value}: Scheduled to new invoice date on {formatted_date}")
            else:
                formatted_date = (date_paid.strftime("%d-%m-%y") if isinstance(date_paid, datetime) else str(date_paid))
                if verbose:
                    print(f"{ws.range(f'{my_dict["Type"]}{row_num}').value}: No invoice to schedule, less than 0 days since {formatted_date}. "
                          f"Next date payment will be: {ws.range(f'{my_dict["Next payment"]}{row_num}').value}")
        else:
            if verbose:
                print(f"{ws.range(f'{my_dict["Type"]}{row_num}').value}: Nothing since no due date or invalid days")
    else:
        if verbose:
            print(f"{ws.range(f'{my_dict["Type"]}{row_num}').value}: Nothing since no payment was made on invoice")

def main():
    run()


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"Failed to open workbook: {e}")

# Emoticons website
# https://emojiterra.com/search/info/#google_vignette
# https://emojidb.org/
