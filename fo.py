import openpyxl as xl
import datetime
from date_variables import date, mnth

fo_share_list = ["ADANI", "APORT", "APOLLO", "AURO", "AXIS", "BAJAJ", "BARODA", "AIRTEL", "BHEL", "BN", "CANBK", "COALIND",
                 "DLF", "DRREDDY", "EICHER", "HCL", "HDFC", "HIND", "IBUL", "ICICI", "INDUSIND", "JIND", "NIFTY", "REL",
                 "SBIN", "TCHEM", "TCON", "TM", "TS", "TCS", "TITAN", "ULTRA", "VEDL"]

fo_daily_aggregate_list = ["ADANI", "AURO", "BN", "CANBK", "DLF", "HIND", "ICICI", "JIND", "NIFTY",
                 "REL", "SBIN", "TCON", "TM", "TS", "TCS", "TITAN"]


foHL_wb = xl.load_workbook(r'C:\Users\admin\PycharmProjects\daily data\fo high low.xlsx')
foHL_sheet = foHL_wb['Sheet1']
foHL_row = 2

fo1_wb = xl.load_workbook(r'C:\Users\admin\PycharmProjects\daily data\fo1.xlsx') # will have to change this when actually doing work on Monday
fo1_sheet = fo1_wb['fo1-Sheet1']
fo1_row = 2

for share in fo_share_list:
    path = rf"E:\Daily Data work\hourlys 1 minute FO\2023\{mnth}\{date}\{share}.xlsx"

    wb = xl.load_workbook(path)
    sheet = wb[f"{share}-Sheet1"]

    start_row = 11
    time_cell = sheet.cell(start_row, 7)

    # loop for changing the time cells format (have to close and reopen otherwise it doesn't change format)
    while time_cell.value is not None:
        time_cell = sheet.cell(start_row, 7)
        time_cell.number_format = 'hh:mm AM/PM'

        start_row += 1

    wb.save(path)

    # reloading the worksheet
    wb = xl.load_workbook(path)
    sheet = wb[f"{share}-Sheet1"]

    start_row = 11

    time_cell = sheet.cell(start_row, 7)
    cur_time = time_cell.value
    end_time = datetime.time(15, 30, 0)

    # 9:25 close value
    cl_9_25 = sheet.cell(start_row, 3).value

    HIGH = 0
    LOW = 9999999

    # HIGH and LOW value finding loop
    while cur_time is not None and cur_time <= end_time:
        time_cell = sheet.cell(start_row, 7)
        high_cell = sheet.cell(start_row, 4)
        low_cell = sheet.cell(start_row, 5)

        cur_time = time_cell.value
        # print(cur_time)

        if high_cell.value is not None and high_cell.value > HIGH:
            HIGH = high_cell.value

        if low_cell.value is not None and low_cell.value < LOW and low_cell.value != 0:
            LOW = low_cell.value

        start_row += 1

    if share in fo_daily_aggregate_list:
        # high
        foHL_sheet.cell(foHL_row, 2).value = HIGH

        # low
        foHL_sheet.cell(foHL_row, 3).value = LOW

        # close
        foHL_sheet.cell(foHL_row, 5).value = fo1_sheet.cell(fo1_row, 5).value

        # vol
        volume = fo1_sheet.cell(fo1_row, 6).value
        volume = str(volume)
        volume = int(volume[0:len(volume)-5])   # truncating volume to display in lakhs

        foHL_sheet.cell(foHL_row, 6).value = volume

        # 9:25 close
        foHL_sheet.cell(foHL_row, 7).value = cl_9_25

        foHL_row += 1
        fo1_row += 1

    print(f"{share} done")


foHL_wb.save(r'C:\Users\admin\PycharmProjects\daily data\fo high low.xlsx')



