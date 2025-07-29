import xlwings as xw
from datetime import datetime, timedelta


def run():
    try:
        wb = xw.Book("house_payment.xlsm")
        ws = wb.sheets["Estimation Expenses"]
        print(f"Opened workbook: {wb.name}")
        for i in range(2, 8):
            check_update(wb, i)

        print(""); print(f"ℹ️ {ws.range('A10').value}")
        for i in range(11, 15):
            check_update(wb, i)

        print(""); print(f"ℹ️ {ws.range('A15').value}")
        for i in range(16, 20):
            check_update(wb, i)

        print(""); print(f"ℹ️ {ws.range('A21').value}")
        for i in range(22, 26):
            check_update(wb, i)

        # wb.save(); wb.close()
        print("xlwings update completed.")
    except Exception as e:
        print(f"Error in run: {e}")


def check_update(wb, row_num):
    check_neutral(wb, row_num)
    update_next_invoice(wb, row_num)


def check_neutral(wb, row_num):
    ws = wb.sheets["Estimation Expenses"]
    status = ws.range(f"L{row_num}").value
    # print(f"Checking row {row_num}, Status: {status}")
    if status == "Neutral":
        due_date = ws.range(f"F{row_num}").value
        print(f"⚠️ {ws.range(f'A{row_num}').value}: Status is Neutral. Due date is {due_date}. Consider payment.")
    elif status == "Bad":
        print(f"📛 {ws.range(f'A{row_num}').value}: Status is Bad.")


def update_next_invoice(wb, row_num):
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
                ws.range(f"G{row_num}").value = ""
                print(f"\U0001F4C5 {ws.range(f'A{row_num}').value}: Scheduled to new invoice date on {ws.range(f'E{row_num}').value}")
        else:
            print(f"\u2705 {ws.range(f'A{row_num}').value}: No invoice to schedule lesser than {limit} "
                  f"days despite recent payment in {date_paid}. Next date payment will be: {ws.range(f'G{row_num}').value}")
    else:
        print(f"{ws.range(f'A{row_num}').value}: Nothing since no payment was made on invoice")


def main():
    run()


if __name__ == "__main__":
    try:
        xw.Book("house_payment.xlsm")
        main()
    except Exception as e:
        print(f"Failed to open workbook: {e}")

# Emoticons website
# https://emojiterra.com/search/info/#google_vignette
