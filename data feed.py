import os

import openpyxl as xl
from openpyxl.styles import Font, Alignment

append = 67  # increment this daily. 67 is for 15-NOV-2023

i = 0  # main iterator variable

# styles
red = Font("Arial", 11, color='ff0000', bold=True)
blue = Font("Arial", 11, color="0000ff", bold=True)
bold = Font("Arial", 11, bold=True)
alignment = Alignment(horizontal='center')

# CASH

cash_shares = {'ADANIENT': 1579, 'APOLLOTYRE': 2946, 'BAJAJFINSERV': 1579, 'BAJAJFINANCE': 1579,
               'BANDHANBANK': 1579, 'BANKBARODA': 1579, 'COAL INDIA': 3232, 'DLF CHL': 4058, 'EICHERMOTOR': 2715,
               'FEDRAL BANK': 1579, 'HCLTECH': 1579, 'HDFC': 3936, 'HINDALCO': 946,'ICICIBANK': 1579, 'INDUSINDBANK': 1579,
               'INFY': 2765, 'JINDALS chl': 5195, 'LICHSGFIN': 1579, 'M&M': 1579, 'M&MFINANCE': 1579,
               'NTPC': 946,'03 RELIANCE CHL': 4793, '04 SBIN CHL': 4860, 'SUNTV': 1579, 'TATACHEM': 1579,
               '07 TATAMOTOR CHL': 4434, 'TATAPOWER': 1579, '05 TATASTEEL chl': 4570, 'ULTRACHEM': 2696}

cash_no_format_list = ['APOLLOTYRE', 'BANDHANBANK', 'BANKBARODA', 'COAL INDIA', 'DLF CHL', '07 TATAMOTOR CHL',
                       '05 TATASTEEL chl', 'TATAPOWER', 'M&MFINANCE', 'FEDRAL BANK', 'HINDALCO', 'NTPC']

# loading 'cash high low.xlsx'
cashHL_wb = xl.load_workbook(r'C:\Users\admin\PycharmProjects\daily data\cash high low.xlsx')
cashHL_sheet = cashHL_wb['Sheet1']
cashHL_row = 2

for share in cash_shares:
    path = rf'E:\Daily Data work\CASH\{share}.xlsx'

    wb = xl.load_workbook(path)
    sheet = wb['D']

    input_row = cash_shares[share] + append  # incrementing row values from base row(start row)

    # data filling
    sheet.cell(input_row, 2).value = cashHL_sheet.cell(cashHL_row, 2).value  # high
    sheet.cell(input_row, 3).value = cashHL_sheet.cell(cashHL_row, 3).value  # low
    sheet.cell(input_row, 4).value = cashHL_sheet.cell(cashHL_row, 4).value  # close
    sheet.cell(input_row, 5).value = cashHL_sheet.cell(cashHL_row, 5).value  # LTP
    sheet.cell(input_row, 6).value = cashHL_sheet.cell(cashHL_row, 6).value  # vol
    sheet.cell(input_row, 7).value = cashHL_sheet.cell(cashHL_row, 7).value  # 9:25 close

    # number formatting
    if share not in cash_no_format_list:
        sheet.cell(input_row, 2).number_format = '0'  # high
        sheet.cell(input_row, 3).number_format = '0'  # low
        sheet.cell(input_row, 4).number_format = '0'  # close
        sheet.cell(input_row, 5).number_format = '0'  # LTP
        sheet.cell(input_row, 7).number_format = '0'  # 9:25 close

    # style formatting
    sheet.cell(input_row, 2).font = blue  # high
    sheet.cell(input_row, 2).alignment = alignment

    sheet.cell(input_row, 3).font = red  # low
    sheet.cell(input_row, 3).alignment = alignment

    sheet.cell(input_row, 4).font = bold  # close
    sheet.cell(input_row, 4).alignment = alignment

    sheet.cell(input_row, 5).font = bold  # LTP
    sheet.cell(input_row, 5).alignment = alignment

    sheet.cell(input_row, 6).font = bold  # vol
    sheet.cell(input_row, 6).alignment = alignment

    sheet.cell(input_row, 7).font = bold  # 9:25 close
    sheet.cell(input_row, 7).alignment = alignment

    i += 1
    cashHL_row += 1

    print(f"{share} done!")

    wb.save(path)

print("----------------------------------CASH DONE----------------------------------")

# FO
i = 0  # main iterator variable

fo_shares = {'ADANI PORT': 2279, 'AUROPHARMA': 2789, '02 BANKNIFTY F': 3309, 'CANBK': 2016, 'DLF': 2831,
             'HINDALCO': 3989, 'ICICIBANK': 1093, 'JINDS': 2274, '01 NIFTY F': 2795, '03 RELIANCE': 2792, 'SBIN': 2793,
             'TATACONSUM': 2276, '05 TATAMOTOR': 2791, '04 TATASTEEL': 2793, 'TCS': 4826, 'TITAN': 1762}

fo_no_format_list = ['04 TATASTEEL']

# loading 'fo high low.xlsx'
foHL_wb = xl.load_workbook(r'C:\Users\admin\PycharmProjects\daily data\fo high low.xlsx')
foHL_sheet = foHL_wb['Sheet1']
foHL_row = 2

for share in fo_shares:
    path = rf'E:\Daily Data work\FO\{share}.xlsx'

    wb = xl.load_workbook(path)
    sheet = wb['D']

    input_row = fo_shares[share] + append

    # data filling
    sheet.cell(input_row, 2).value = foHL_sheet.cell(foHL_row, 2).value  # high
    sheet.cell(input_row, 3).value = foHL_sheet.cell(foHL_row, 3).value  # low
    sheet.cell(input_row, 4).value = foHL_sheet.cell(foHL_row, 4).value  # close
    sheet.cell(input_row, 5).value = foHL_sheet.cell(foHL_row, 5).value  # LTP
    sheet.cell(input_row, 6).value = foHL_sheet.cell(foHL_row, 6).value  # vol
    sheet.cell(input_row, 7).value = foHL_sheet.cell(foHL_row, 7).value  # 9:25 close

    # number formatting
    if share not in cash_no_format_list:
        sheet.cell(input_row, 2).number_format = '0'  # high
        sheet.cell(input_row, 3).number_format = '0'  # low
        sheet.cell(input_row, 4).number_format = '0'  # close
        sheet.cell(input_row, 5).number_format = '0'  # LTP
        sheet.cell(input_row, 7).number_format = '0'  # 9:25 close

    # style formatting
    sheet.cell(input_row, 2).font = blue  # high
    sheet.cell(input_row, 2).alignment = alignment

    sheet.cell(input_row, 3).font = red  # low
    sheet.cell(input_row, 3).alignment = alignment

    sheet.cell(input_row, 4).font = bold  # close
    sheet.cell(input_row, 4).alignment = alignment

    sheet.cell(input_row, 5).font = bold  # LTP
    sheet.cell(input_row, 5).alignment = alignment

    sheet.cell(input_row, 6).font = bold  # vol
    sheet.cell(input_row, 6).alignment = alignment

    sheet.cell(input_row, 7).font = bold  # 9:25 close
    sheet.cell(input_row, 7).alignment = alignment

    i += 1
    foHL_row += 1

    print(f"{share} done!")

    wb.save(path)

print("----------------------------------FO DONE----------------------------------")

try:
    os.remove(r'E:\Daily Data work\csh.xls')
    os.remove(r'E:\Daily Data work\fo1.xls')
except FileNotFoundError:
    pass

