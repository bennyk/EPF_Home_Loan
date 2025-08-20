import xlwings as xw
from datetime import datetime, timedelta

verbose = 0


def run():
    try:
        wb = xw.Book("house_payment.xlsm")
        ws = wb.sheets["Estimation Expenses"]

        # TODO xlwings sheets does not implement protect or unprotect sheets. Hmm
        # ws.unprotect()
        # ws.protect(password="xlwings")

        print(f"Opened workbook: {wb.name}")
        for i in range(2, 11):
            check_update(wb, i)

        for i in range(12, 33):
            if ws.range(f"A{i}").value[0] == "🏠":
                print(f"ℹ️ {ws.range(f'A{i}').value}")
            else:
                check_update(wb, i)

        # wb.save(); wb.close()
        print("xlwings update completed.")
    except Exception as e:
        print(f"Error in run: {e}")


def check_update(wb, row_num):
    # Update first before checking.
    update_next_invoice(wb, row_num)
    check_neutral(wb, row_num)


def check_neutral(wb, row_num):
    ws = wb.sheets["Estimation Expenses"]
    status = ws.range(f"L{row_num}").value
    item_name = ws.range(f"A{row_num}").value
    # print(f"Checking row {row_num}, Status: {status}")

    # Target "Date Paid" cell (Column J)
    date_paid_cell = ws.range(f"J{row_num}")
    date_paid_cell.color = None

    if status == "Neutral":
        due_date = ws.range(f"F{row_num}").value
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
    global verbose
    ws = wb.sheets["Estimation Expenses"]
    date_paid = ws.range(f"J{row_num}").value
    # print(f"Processing row {row_num}, Date Paid: {date_paid}")
    if date_paid is not None and date_paid != "":
        if isinstance(date_paid, (int, float)):
            date_paid = datetime(1899, 12, 30) + timedelta(days=date_paid)
        limit = 20
        if (datetime.now() - date_paid).days > limit:
            next_payment = ws.range(f"G{row_num}").value
            if next_payment != "Nope":
                ws.range(f"E{row_num}").value = next_payment.strftime("%d-%m-%y") if isinstance(next_payment, datetime) else next_payment
                ws.range(f"I{row_num}").value = ws.range(f"I{row_num}").value + 1
                ws.range(f"G{row_num}").value = "Nope"
                ws.range(f"J{row_num}").value = ""
                print(f"\U0001F4C5 {ws.range(f'A{row_num}').value}: "
                      f"Scheduled to new invoice date on {next_payment.strftime('%d-%m-%y')}")
        else:
            if verbose > 0:
                print(f"\u2705 {ws.range(f'A{row_num}').value}: No invoice to schedule lesser than {limit} "
                      f"days despite recent payment in {date_paid.strftime("%d-%m-%y")}. Next date payment will be: "
                      f"{ws.range(f'G{row_num}').value.strftime("%d-%m-%y")}")
    else:
        if verbose > 0:
            print(f"{ws.range(f'A{row_num}').value}: Nothing since no payment was made on invoice")


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
