import os
import shutil

import openpyxl as xl
from openpyxl.styles import PatternFill
import datetime
from xls2xlsx import XLS2XLSX
from date_variables import date, mnth, yr

fo_share_list = ["ADANI", "APORT", "APOLLO", "AURO", "AXIS", "BAJAJ", "BARODA", "AIRTEL", "BHEL", "BN", "CANBK", "COALIND",
                 "DLF", "DRREDDY", "EICHER", "HCL", "HDFC", "HIND", "IBUL", "ICICI", "INDUSIND", "JIND", "NIFTY", "REL",
                 "SBIN", "TCHEM", "TCON", "TM", "TS", "TCS", "TITAN", "ULTRA", "VEDL"]
# fo_share_list = ["ADANI"]

fo_daily_aggregate_list = ["APORT", "AURO", "BN", "CANBK", "DLF", "HIND", "ICICI", "JIND", "NIFTY", "REL", "SBIN",
                           "TCON", "TM", "TS", "TCS", "TITAN"]


foHL_wb = xl.load_workbook(r'C:\Users\admin\PycharmProjects\daily data\fo high low.xlsx')
foHL_sheet = foHL_wb['Sheet1']
foHL_row = 2

fo1_wb = xl.load_workbook(r'E:\Daily Data work\fo1.xlsx')
fo1_sheet = fo1_wb['fo1-Sheet1']
fo1_row = 2

fl_9_25 = 0.392361111       # 9:25 time value in general format

# copying hourlys (.xls) as backup
# path to source directory
src_dir = rf"E:\Daily Data work\hourlys 1 minute FO\{yr}\{mnth}\{date}"

# path to destination directory
dest_dir = rf"C:\Users\admin\PycharmProjects\daily data\Daily Backup hourlys\1 min fo"

# getting all the files in the source directory
files = os.listdir(src_dir)
shutil.copytree(src_dir, dest_dir)
print("Files copied as backup!")

for share in fo_share_list:
    path = rf"E:\Daily Data work\hourlys 1 minute FO\{yr}\{mnth}\{date}\{share}.xlsx"
    xls_path = rf"E:\Daily Data work\hourlys 1 minute FO\{yr}\{mnth}\{date}\{share}.xls"
    x2x = XLS2XLSX(xls_path)

    wb = x2x.to_xlsx()
    sheet = wb[f"{share}-Sheet1"]

    start_row = 2
    time_cell = sheet.cell(start_row, 7)

    while time_cell.value < fl_9_25:
        start_row += 1
        time_cell = sheet.cell(start_row, 7)

    print(f"starting row is {start_row}")

    start_row_2 = start_row

    # loop for changing the time cells format (have to close and reopen otherwise it doesn't change format)
    while time_cell.value is not None:
        time_cell = sheet.cell(start_row, 7)
        time_cell.number_format = 'hh:mm AM/PM'

        start_row += 1

    wb.save(path)

    # reloading the worksheet
    wb = xl.load_workbook(path)
    sheet = wb[f"{share}-Sheet1"]

    start_row = start_row_2

    time_cell = sheet.cell(start_row, 7)
    cur_time = time_cell.value
    end_time = datetime.time(15, 30, 0)

    # 9:25 close value
    cl_9_25 = sheet.cell(start_row, 3).value
    sheet.cell(start_row, 3).fill = PatternFill("solid", 'FFFF00')

    # reloading wb otherwise pattern fill doesn't work
    wb.save(path)
    wb = xl.load_workbook(path)
    sheet = wb[f"{share}-Sheet1"]

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
        foHL_sheet.cell(foHL_row, 4).value = fo1_sheet.cell(fo1_row, 5).value

        # vol
        volume = fo1_sheet.cell(fo1_row, 6).value
        volume = str(volume)
        volume = int(volume[0:len(volume)-5])   # truncating volume to display in lakhs

        foHL_sheet.cell(foHL_row, 6).value = volume

        # 9:25 close
        foHL_sheet.cell(foHL_row, 7).value = cl_9_25

        foHL_row += 1
        fo1_row += 1

    HIGH = 0
    LOW = 9999999

    # 30 MIN FORMATTING IN 1 MIN SHEETS
    sheet.cell(1, 14).value = "HIGH"
    sheet.cell(1, 15).value = "LOW"
    sheet.cell(1, 16).value = "CLOSE"

    start_row = start_row_2     # actual start row

    time_cell = sheet.cell(start_row, 7)
    cur_time = time_cell.value

    count = 0

    while cur_time is not None and cur_time <= end_time:
        high_cell = sheet.cell(start_row, 4)
        low_cell = sheet.cell(start_row, 5)

        # print(cur_time)

        if high_cell.value is not None and high_cell.value > HIGH:
            HIGH = high_cell.value

        if low_cell.value is not None and low_cell.value < LOW and low_cell.value != 0:
            LOW = low_cell.value

        # resetting after 30 mins
        if count == 30:
            sheet.cell(start_row, 14).value = HIGH
            sheet.cell(start_row, 15).value = LOW

            # if 30 min close is empty or 0
            if sheet.cell(start_row, 3).value == 0 or sheet.cell(start_row, 3).value is None:
                temp_row = start_row

                while sheet.cell(temp_row, 3).value == 0 or sheet.cell(temp_row, 3).value is None:
                    temp_row -= 1

                sheet.cell(start_row, 16).value = sheet.cell(temp_row, 3).value  # close

            else:
                sheet.cell(start_row, 16).value = sheet.cell(start_row, 3).value  # close

            count = 1
            HIGH = 0
            LOW = 9999999
            start_row += 1
            continue

        start_row += 1
        count += 1

        time_cell = sheet.cell(start_row, 7)
        cur_time = time_cell.value

    # last any left aggregate (< 30 mins)
    sheet.cell(start_row-1, 14).value = HIGH
    sheet.cell(start_row-1, 15).value = LOW
    sheet.cell(start_row-1, 16).value = sheet.cell(start_row-1, 3).value  # close

    wb.save(path)
    os.remove(xls_path)
    # wb.save(rf'E:\Daily Data work\hourlys 1 minute FO\{yr}\{mnth}\{date}\{share}1.xlsx')
    print(f"{share} done")


# MD file
md_path = rf"E:\Daily Data work\MD files\{yr}\{mnth}\fo{date[:2]}{mnth}20{date[6:]}bhav.xlsx"
md_wb = xl.load_workbook(md_path)

md_sheet = md_wb[rf"fo{date[:2]}{mnth}20{date[6:]}bhav"]

k = 0       # variable to iterate over all shares in share_list and index_list
md_row = 2  # starting row of md file

# lists for shares as per their names in md file and their respective indices in the 'fo high low.xlsx' sheet
share_list = ["BANKNIFTY", "NIFTY", "ADANIPORTS", "AUROPHARMA", "CANBK", "DLF", "HINDALCO",
              "ICICIBANK", "JINDALSTEL", "RELIANCE", "SBIN", "TATACONSUM", "TATAMOTORS",
              "TATASTEEL", "TCS", "TITAN"]

index_list = [4, 10, 2, 3, 5, 6, 7, 8, 9, 11, 12, 13, 14, 15, 16, 17]

while k < len(share_list):
    # print(md_row)

    name_cell = md_sheet.cell(md_row, 2)

    if name_cell.value == share_list[k]:
        foHL_sheet.cell(index_list[k], 5).value = md_sheet.cell(md_row, 10).value

        k += 1
        md_row += 2     # to skip the next 2 expires of the share

    md_row += 1

foHL_wb.save(r'C:\Users\admin\PycharmProjects\daily data\fo high low.xlsx')
