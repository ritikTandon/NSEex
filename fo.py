import openpyxl as xl
import datetime
from date_variables import date, mnth

fo_share_list = ["ADANI", "APORT", "APOLLO", "AURO", "AXIS", "BAJAJ", "BARODA", "AIRTEL", "BHEL", "BN", "CANBK", "COALIND",
                 "DLF", "DRREDDY", "EICHER", "HCL", "HDFC", "HIND", "IBUL", "ICICI", "INDUSIND", "JIND", "NIFTY", "REL",
                 "SBIN", "TCHEM", "TCON", "TM", "TS", "TCS", "TITAN", "ULTRA", "VEDL"]

fo_daily_aggregate_list = ["APORT", "AURO", "BN", "CANBK", "DLF", "HIND", "ICICI", "JIND", "NIFTY",
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

    print(f"{share} done")


md_path = rf"E:\Daily Data work\MD files\2023\{mnth}\fo{date[:2]}{mnth}20{date[6:]}bhav.xlsx"
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



