import openpyxl as xl
from openpyxl.styles import Font, Alignment, PatternFill
import datetime
from date_variables import date, mnth, yr

append = 2       # increment this daily. 2 is for 16-AUG-2023

i = 0       # main iterator variable

# styles
red = Font("Arial", 11, color='ff0000', bold=True)
blue = Font("Arial", 11, color="0000ff", bold=True)
bold = Font("Arial", 11, bold=True)
alignment = Alignment(horizontal='center')

# CASH
cash_share_names = ['ADANIENT', 'APOLLOTYRE', 'BAJAJFINSERV', 'BAJAJFINANCE', 'BANDHANBANK', 'BANKBARODA', 'COAL INDIA',
                    'DLF CHL', 'EICHERMOTOR', 'FEDRAL BANK', 'HCLTECH', 'HDFC', 'ICICIBANK', 'INDUSINDBANK', 'INFY',
                    'JINDALS chl', 'LICHSGFIN', 'M&M', 'M&MFINANCE', '03 RELIANCE CHL', '04 SBIN CHL', 'SUNTV',
                    'TATACHEM', '07 TATAMOTOR CHL', 'TATAPOWER', '05 TATASTEEL chl', 'ULTRACHEM']

cash_share_rows = [1579, 2946, 1579, 1579, 1579, 1579, 3232, 4058, 2715, 1579, 1579, 3936, 1579, 1579, 2765, 5195, 1579,
                   1579, 1579, 4793, 4860, 1579, 1579, 4433, 1579, 4570, 2696]

cash_no_format_list = ['APOLLOTYRE', 'BANDHANBANK', 'BANKBARODA', 'COAL INDIA', 'DLF CHL', '07 TATAMOTOR CHL',
                       '05 TATASTEEL chl', 'TATAPOWER', 'M&MFINANCE', 'FEDRAL BANK']

# loading 'cash high low.xlsx'
cashHL_wb = xl.load_workbook(r'C:\Users\admin\PycharmProjects\daily data\cash high low.xlsx')
cashHL_sheet = cashHL_wb['Sheet1']
cashHL_row = 2

for share in cash_share_names:
    if share == "ICICIBANK":
        cashHL_row += 1

    path = rf'E:\Daily Data work\CASH\{share}.xlsx'

    wb = xl.load_workbook(path)
    sheet = wb['D']

    input_row = cash_share_rows[i]

    # data filling
    sheet.cell(cash_share_rows[i], 2).value = cashHL_sheet.cell(cashHL_row, 2).value    # high
    sheet.cell(cash_share_rows[i], 3).value = cashHL_sheet.cell(cashHL_row, 3).value    # low
    sheet.cell(cash_share_rows[i], 4).value = cashHL_sheet.cell(cashHL_row, 4).value    # close
    sheet.cell(cash_share_rows[i], 5).value = cashHL_sheet.cell(cashHL_row, 5).value    # LTP
    sheet.cell(cash_share_rows[i], 6).value = cashHL_sheet.cell(cashHL_row, 6).value    # vol
    sheet.cell(cash_share_rows[i], 7).value = cashHL_sheet.cell(cashHL_row, 7).value    # 9:25 close

    # number formatting
    if share not in cash_no_format_list:
        sheet.cell(cash_share_rows[i], 2).number_format = '0'  # high
        sheet.cell(cash_share_rows[i], 3).number_format = '0'  # low
        sheet.cell(cash_share_rows[i], 4).number_format = '0'  # close
        sheet.cell(cash_share_rows[i], 5).number_format = '0'  # LTP
        sheet.cell(cash_share_rows[i], 7).number_format = '0'  # 9:25 close

    # style formatting
    sheet.cell(cash_share_rows[i], 2).font = blue   # high
    sheet.cell(cash_share_rows[i], 2).alignment = alignment

    sheet.cell(cash_share_rows[i], 3).font = red  # low
    sheet.cell(cash_share_rows[i], 3).alignment = alignment

    sheet.cell(cash_share_rows[i], 4).font = bold  # close
    sheet.cell(cash_share_rows[i], 4).alignment = alignment

    sheet.cell(cash_share_rows[i], 5).font = bold  # LTP
    sheet.cell(cash_share_rows[i], 5).alignment = alignment

    sheet.cell(cash_share_rows[i], 6).font = bold  # vol
    sheet.cell(cash_share_rows[i], 6).alignment = alignment

    sheet.cell(cash_share_rows[i], 7).font = bold  # 9:25 close
    sheet.cell(cash_share_rows[i], 7).alignment = alignment

    i += 1
    cashHL_row += 1

    print(f"{share} done!")

    wb.save(path)

print("----------------------------------CASH DONE----------------------------------")

# FO
i = 0   # main iterator variable

fo_share_names = ['ADANI PORT', 'AUROPHARMA', '02 BANKNIFTY F', 'CANBK', 'DLF', 'HINDALCO', 'ICICIBANK', 'JINDS',
                  '01 NIFTY F', '03 RELIANCE', 'SBIN', 'TATACONSUM', '05 TATAMOTOR', '04 TATASTEEL', 'TCS', 'TITAN']

fo_share_rows = [2279, 2789, 3309, 2016, 2831, 3989, 1093, 2274, 2795, 2792, 2793, 2276, 2791, 2793, 4826, 1762]

fo_no_format_list = ['04 TATASTEEL']

# loading 'cash high low.xlsx'
foHL_wb = xl.load_workbook(r'C:\Users\admin\PycharmProjects\daily data\fo high low.xlsx')
foHL_sheet = foHL_wb['Sheet1']
foHL_row = 2

for share in fo_share_names:
    path = rf'E:\Daily Data work\FO\{share}.xlsx'

    wb = xl.load_workbook(path)
    sheet = wb['D']

    input_row = fo_share_rows[i]

    # data filling
    sheet.cell(fo_share_rows[i], 2).value = foHL_sheet.cell(foHL_row, 2).value    # high
    sheet.cell(fo_share_rows[i], 3).value = foHL_sheet.cell(foHL_row, 3).value    # low
    sheet.cell(fo_share_rows[i], 4).value = foHL_sheet.cell(foHL_row, 4).value    # close
    sheet.cell(fo_share_rows[i], 5).value = foHL_sheet.cell(foHL_row, 5).value    # LTP
    sheet.cell(fo_share_rows[i], 6).value = foHL_sheet.cell(foHL_row, 6).value    # vol
    sheet.cell(fo_share_rows[i], 7).value = foHL_sheet.cell(foHL_row, 7).value    # 9:25 close

    # number formatting
    if share not in cash_no_format_list:
        sheet.cell(fo_share_rows[i], 2).number_format = '0'  # high
        sheet.cell(fo_share_rows[i], 3).number_format = '0'  # low
        sheet.cell(fo_share_rows[i], 4).number_format = '0'  # close
        sheet.cell(fo_share_rows[i], 5).number_format = '0'  # LTP
        sheet.cell(fo_share_rows[i], 7).number_format = '0'  # 9:25 close

    # style formatting
    sheet.cell(fo_share_rows[i], 2).font = blue   # high
    sheet.cell(fo_share_rows[i], 2).alignment = alignment

    sheet.cell(fo_share_rows[i], 3).font = red  # low
    sheet.cell(fo_share_rows[i], 3).alignment = alignment

    sheet.cell(fo_share_rows[i], 4).font = bold  # close
    sheet.cell(fo_share_rows[i], 4).alignment = alignment

    sheet.cell(fo_share_rows[i], 5).font = bold  # LTP
    sheet.cell(fo_share_rows[i], 5).alignment = alignment

    sheet.cell(fo_share_rows[i], 6).font = bold  # vol
    sheet.cell(fo_share_rows[i], 6).alignment = alignment

    sheet.cell(fo_share_rows[i], 7).font = bold  # 9:25 close
    sheet.cell(fo_share_rows[i], 7).alignment = alignment

    i += 1
    foHL_row += 1

    print(f"{share} done!")

    wb.save(path)

print("----------------------------------FO DONE----------------------------------")


