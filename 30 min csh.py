import os
import shutil

import openpyxl as xl
from openpyxl.styles import Font, Alignment
from xls2xlsx import XLS2XLSX

from date_variables import date, mnth, yr

cashHL_wb = xl.load_workbook(r'C:\Users\admin\PycharmProjects\daily data\cash high low.xlsx')
cashHL_sheet = cashHL_wb['Sheet1']

foHL_wb = xl.load_workbook(r'C:\Users\admin\PycharmProjects\daily data\fo high low.xlsx')
foHL_sheet = foHL_wb['Sheet1']

algoHL_wb = xl.load_workbook(r'C:\Users\admin\PycharmProjects\daily data\algo high low.xlsx')
algoHL_sheet = algoHL_wb['Sheet1']

# dict to store share names with their row in 'cash/algo/fo high low.xlsx' sheets respectively
cash_30_min_list = {"AARTIIND": 2, "ABB": 3, "ADANI": 3, "APOLLO": 4, "ASHOKLEY": 12, "BAJFINSV": 5, "BAJFIN": 6, "BANBK": 7,
                    "BHEL": 27, "BARODA": 8, "BN": 2, "CHAMBAL": 33, "COALIND": 9, "DIXON": 47, "DLF": 10, "EICHER": 11, "ESCORTS": 50, "FEDBANK": 12, "HCL": 13,
                    "HINDALCO": 15, "IGL": 70, "INDUSIND": 17, "JIND": 19, "LIC": 20, "M&M": 21, "M&MFIN": 22,
                    "NIFTY": 3, "NTPC": 23, "ONGC": 106, "RECLTD": 116, "REL": 24, "SBIN": 25, "SUNTV": 26, "TM": 28,
                    "TP": 29, "TS": 30, "VEDL": 136}

# dict to store share names with their row in 'cash/algo/fo high low.xlsx' sheets respectively
cash_15_min_list = {"ABB": 3, "APOLLO": 4, "BAJFINSV": 5, "BAJFIN": 6, "BHEL": 27, "BSOFT": 30, "CHAMBAL": 33, "COFORGE": 36,
                    "DIXON": 47, "DLF": 10, "GLENMARK": 52, "HAL": 60, "LAURUSLABS": 85, "MCX": 94, "NIFTY": 3,
                    "REL": 24, "TM": 28}

# no decimal points in display

format_list = ["NIFTY", "EICHER", "BN", "DIXON", "ABB"]

# shares that get their data from 'algo high low.xlsx', when adding shares that are in algo and not in cash, add to this
algo_shares = ["ESCORTS", "IGL", "VEDL", "ABB", "ASHOKLEY", "DIXON", "ONGC", "RECLTD", "BHEL", "CHAMBAL"]

# 15 min shares that get their data from 'algo high low.xlsx', when adding shares that are in algo and not in cash, add to this
algo_shares_15_min = ["ABB", "DIXON", "BSOFT", "BHEL", "CHAMBAL", "COFORGE", "GLENMARK", "HAL", "LAURUSLABS", "MCX"]

# copying 30 min hourlys (.xls) as backup
# path to source directory
src_dir = rf"E:\Daily Data work\hourlys 30 minute CASH\{yr}\{mnth}\{date}"

# path to destination directory
dest_dir = rf"C:\Users\admin\PycharmProjects\daily data\Daily Backup hourlys\30 min csh"

# getting all the files in the source directory
src_files = os.listdir(src_dir)
for file_name in src_files:
    full_file_name = os.path.join(src_dir, file_name)
    if os.path.isfile(full_file_name):
        shutil.copy(full_file_name, dest_dir)

print("30 min Files copied as backup!")


# copying 15 min hourlys (.xls) as backup
# path to source directory
src_dir = rf"E:\Daily Data work\hourlys 15 minute CASH\{yr}\{mnth}\{date}"

# path to destination directory
dest_dir = rf"C:\Users\admin\PycharmProjects\daily data\Daily Backup hourlys\15 min csh"

# getting all the files in the source directory
src_files = os.listdir(src_dir)
for file_name in src_files:
    full_file_name = os.path.join(src_dir, file_name)
    if os.path.isfile(full_file_name):
        shutil.copy(full_file_name, dest_dir)

print("15 min Files copied as backup!")


# LTP and PREV
# 30 min
ltp_wb = xl.load_workbook(rf'C:\Users\admin\PycharmProjects\daily data\LTP PREV.xlsx')
ltp_sheet_30 = ltp_wb["30"]
ltp_row = 2

while ltp_row <= len(ltp_sheet_30["A"]):
    ltp_sheet_30.cell(ltp_row, 3).value = ltp_sheet_30.cell(ltp_row, 2).value  # moving last day's LTP to 'PREV'

    share_name = ltp_sheet_30.cell(ltp_row, 1).value

    if share_name in ["BN", "NIFTY"]:    # separate for NIFTY and BN as their data LTP come from 'fo high low.xlsx'
        ltp_sheet_30.cell(ltp_row, 2).value = foHL_sheet.cell(cash_30_min_list[share_name], 5).value

    elif share_name in algo_shares:  # separate for shares who's LTP come from 'algo high low.xlsx'
        ltp_sheet_30.cell(ltp_row, 2).value = algoHL_sheet.cell(cash_30_min_list[share_name], 5).value
    else:
        ltp_sheet_30.cell(ltp_row, 2).value = cashHL_sheet.cell(cash_30_min_list[share_name], 5).value

    ltp_row += 1

# 15 min
ltp_sheet_15 = ltp_wb["15"]
ltp_row = 2

while ltp_row <= len(ltp_sheet_15["A"]):
    ltp_sheet_15.cell(ltp_row, 3).value = ltp_sheet_15.cell(ltp_row, 2).value  # moving last day's LTP to 'PREV'

    share_name = ltp_sheet_15.cell(ltp_row, 1).value

    if share_name in ["BN", "NIFTY"]:    # separate for NIFTY and BN as their data LTP come from 'fo high low.xlsx'
        ltp_sheet_15.cell(ltp_row, 2).value = foHL_sheet.cell(cash_30_min_list[share_name], 5).value

    elif share_name in algo_shares:  # separate for shares who's LTP come from 'algo high low.xlsx'
        ltp_sheet_15.cell(ltp_row, 2).value = algoHL_sheet.cell(cash_30_min_list[share_name], 5).value
    else:
        ltp_sheet_15.cell(ltp_row, 2).value = cashHL_sheet.cell(cash_30_min_list[share_name], 5).value

    ltp_row += 1

ltp_row = 2

for share in cash_30_min_list:
    path = rf"E:\Daily Data work\hourlys 30 minute CASH\{yr}\{mnth}\{date}\{share}.xlsx"
    xls_path = rf"E:\Daily Data work\hourlys 30 minute CASH\{yr}\{mnth}\{date}\{share}.xls"
    x2x = XLS2XLSX(xls_path)

    cash_30_min_wb = x2x.to_xlsx()
    old_30_min_sheet = cash_30_min_wb[f"{share}-Sheet1"]

    new_30_min_sheet = cash_30_min_wb.create_sheet(f"{share}")

    # FIXED HEADINGS
    new_30_min_sheet.cell(6, 6).value = f'{share}'
    new_30_min_sheet.cell(6, 7).value = "HIGH"
    new_30_min_sheet.cell(6, 8).value = "LOW"
    new_30_min_sheet.cell(6, 9).value = "LTP"
    new_30_min_sheet.cell(6, 10).value = "PREV"

    new_30_min_sheet.cell(8, 6).value = "Time"
    new_30_min_sheet.cell(8, 7).value = "High Rate"
    new_30_min_sheet.cell(8, 8).value = "Low Rate"
    new_30_min_sheet.cell(8, 9).value = "Close Rate"

    old_sheet_row = 2

    while old_sheet_row <= len(old_30_min_sheet["A"]):
        # time
        new_30_min_sheet.cell(old_sheet_row + 7, 6).value = old_30_min_sheet.cell(old_sheet_row, 7).value
        new_30_min_sheet.cell(old_sheet_row + 7, 6).number_format = 'hh:mm AM/PM'

        # high
        new_30_min_sheet.cell(old_sheet_row + 7, 7).value = old_30_min_sheet.cell(old_sheet_row, 4).value

        # low
        new_30_min_sheet.cell(old_sheet_row + 7, 8).value = old_30_min_sheet.cell(old_sheet_row, 5).value

        # close
        new_30_min_sheet.cell(old_sheet_row + 7, 9).value = old_30_min_sheet.cell(old_sheet_row, 3).value

        old_sheet_row += 1

    del cash_30_min_wb[f"{share}-Sheet1"]

    # bolding the sheet
    for i in range(25):
        for j in range(15):
            new_30_min_sheet.cell(i+1, j+1).font = Font(bold=True)
            new_30_min_sheet.cell(i+1, j+1).alignment = Alignment(horizontal='center')

            # formatting to 0 decimal places if they are in format_list
            if share in format_list:
                if j > 5:
                    new_30_min_sheet.cell(i+1, j+1).number_format = '0'

    new_30_min_sheet.cell(7, 6).number_format = '0'     # for the 9:25 cl formatting

    # deleting 4:00 pm row
    new_30_min_sheet.delete_rows(22, 1)

    # filling LTP and PREV
    new_30_min_sheet.cell(7, 9).value = ltp_sheet_30.cell(ltp_row, 2).value  # LTP
    new_30_min_sheet.cell(7, 10).value = ltp_sheet_30.cell(ltp_row, 3).value  # PREV
    ltp_row += 1

    # filling rest of the data
    if share in ["BN", "NIFTY"]:    # separate for NIFTY and BN as their data will come from 'fo high low.xlsx'
        new_30_min_sheet.cell(7, 6).value = foHL_sheet.cell(cash_30_min_list[share], 7).value   # 9:25 cl
        new_30_min_sheet.cell(7, 7).value = foHL_sheet.cell(cash_30_min_list[share], 2).value   # HIGH
        new_30_min_sheet.cell(7, 8).value = foHL_sheet.cell(cash_30_min_list[share], 3).value   # LOW

    elif share in algo_shares:
        new_30_min_sheet.cell(7, 6).value = algoHL_sheet.cell(cash_30_min_list[share], 7).value  # 9:25 cl
        new_30_min_sheet.cell(7, 7).value = algoHL_sheet.cell(cash_30_min_list[share], 2).value  # HIGH
        new_30_min_sheet.cell(7, 8).value = algoHL_sheet.cell(cash_30_min_list[share], 3).value  # LOW

    else:
        new_30_min_sheet.cell(7, 6).value = cashHL_sheet.cell(cash_30_min_list[share], 7).value  # 9:25 cl
        new_30_min_sheet.cell(7, 7).value = cashHL_sheet.cell(cash_30_min_list[share], 2).value  # HIGH
        new_30_min_sheet.cell(7, 8).value = cashHL_sheet.cell(cash_30_min_list[share], 3).value  # LOW

    cash_30_min_wb.save(path)
    os.remove(xls_path)

ltp_row = 2
# 15 min shares
for share in cash_15_min_list:
    path = rf"E:\Daily Data work\hourlys 15 minute CASH\{yr}\{mnth}\{date}\{share}.xlsx"
    xls_path = rf"E:\Daily Data work\hourlys 15 minute CASH\{yr}\{mnth}\{date}\{share}.xls"
    x2x = XLS2XLSX(xls_path)

    cash_15_min_wb = x2x.to_xlsx()
    old_15_min_sheet = cash_15_min_wb[f"{share}-Sheet1"]

    new_15_min_sheet = cash_15_min_wb.create_sheet(f"{share}")

    # FIXED HEADINGS
    new_15_min_sheet.cell(6, 6).value = f'{share}'
    new_15_min_sheet.cell(6, 7).value = "HIGH"
    new_15_min_sheet.cell(6, 8).value = "LOW"
    new_15_min_sheet.cell(6, 9).value = "LTP"
    new_15_min_sheet.cell(6, 10).value = "PREV"

    new_15_min_sheet.cell(8, 6).value = "Time"
    new_15_min_sheet.cell(8, 7).value = "High Rate"
    new_15_min_sheet.cell(8, 8).value = "Low Rate"
    new_15_min_sheet.cell(8, 9).value = "Close Rate"

    old_sheet_row = 2

    while old_sheet_row <= len(old_15_min_sheet["A"]):
        # time
        new_15_min_sheet.cell(old_sheet_row + 7, 6).value = old_15_min_sheet.cell(old_sheet_row, 7).value
        new_15_min_sheet.cell(old_sheet_row + 7, 6).number_format = 'hh:mm AM/PM'

        # high
        new_15_min_sheet.cell(old_sheet_row + 7, 7).value = old_15_min_sheet.cell(old_sheet_row, 4).value

        # low
        new_15_min_sheet.cell(old_sheet_row + 7, 8).value = old_15_min_sheet.cell(old_sheet_row, 5).value

        # close
        new_15_min_sheet.cell(old_sheet_row + 7, 9).value = old_15_min_sheet.cell(old_sheet_row, 3).value

        old_sheet_row += 1

    del cash_15_min_wb[f"{share}-Sheet1"]

    # bolding the sheet
    for i in range(60):
        for j in range(40):
            new_15_min_sheet.cell(i+1, j+1).font = Font(bold=True)
            new_15_min_sheet.cell(i+1, j+1).alignment = Alignment(horizontal='center')

            # formatting to 0 decimal places if they are in format_list
            if share in format_list:
                if j > 5:
                    new_15_min_sheet.cell(i+1, j+1).number_format = '0'

    new_15_min_sheet.cell(7, 6).number_format = '0'     # for the 9:25 cl formatting

    # deleting 4:00 pm row
    # new_15_min_sheet.delete_rows(22, 1) # check if this is needed

    # filling LTP and PREV
    new_15_min_sheet.cell(7, 9).value = ltp_sheet_15.cell(ltp_row, 2).value  # LTP
    new_15_min_sheet.cell(7, 10).value = ltp_sheet_15.cell(ltp_row, 3).value  # PREV
    ltp_row += 1

    # filling rest of the data
    if share in ["BN", "NIFTY"]:    # separate for NIFTY and BN as their data will come from 'fo high low.xlsx'
        new_15_min_sheet.cell(7, 6).value = foHL_sheet.cell(cash_15_min_list[share], 7).value   # 9:25 cl
        new_15_min_sheet.cell(7, 7).value = foHL_sheet.cell(cash_15_min_list[share], 2).value   # HIGH
        new_15_min_sheet.cell(7, 8).value = foHL_sheet.cell(cash_15_min_list[share], 3).value   # LOW

    elif share in algo_shares:
        new_15_min_sheet.cell(7, 6).value = algoHL_sheet.cell(cash_15_min_list[share], 7).value  # 9:25 cl
        new_15_min_sheet.cell(7, 7).value = algoHL_sheet.cell(cash_15_min_list[share], 2).value  # HIGH
        new_15_min_sheet.cell(7, 8).value = algoHL_sheet.cell(cash_15_min_list[share], 3).value  # LOW

    else:
        new_15_min_sheet.cell(7, 6).value = cashHL_sheet.cell(cash_15_min_list[share], 7).value  # 9:25 cl
        new_15_min_sheet.cell(7, 7).value = cashHL_sheet.cell(cash_15_min_list[share], 2).value  # HIGH
        new_15_min_sheet.cell(7, 8).value = cashHL_sheet.cell(cash_15_min_list[share], 3).value  # LOW

    cash_15_min_wb.save(path)
    os.remove(xls_path)

ltp_wb.save(rf'C:\Users\admin\PycharmProjects\daily data\LTP PREV.xlsx')
