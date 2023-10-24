import openpyxl as xl
import pandas as pd
from win32com.client import Dispatch
from openpyxl.styles import Font, Alignment
from date_variables import date, mnth, yr
from zipfile import ZipFile
from decimal import Decimal, ROUND_UP


# up and down rounding functions (to nearest 0.05)
def round_up(var):
    x = Decimal(str(var))

    return float((x * 2).quantize(Decimal('.5'), rounding=ROUND_UP) / 2)


def round_down(var):
    x = Decimal(str(var))

    return float((x * 2).quantize(Decimal('.5')) / 2)


# reading previous dates from date.txt
with open("date.txt", "r") as file:
    data = file.readlines()
    old_date = rf"{data[0][:-1]}"
    old_mnth = rf"{data[1][:-1]}"
    old_yr = rf"{data[2]}"

md_path_zipped = rf"E:\chrome downloads\cm{date[:2]}{mnth}20{date[6:]}bhav.csv.zip"     # .zip file path of downloaded cash bhavcpoy
md_path = rf"E:\chrome downloads"

# extracting .zip file
with ZipFile(md_path_zipped, 'r') as zObject:
    zObject.extractall(path=md_path)

# list of shares we want to check
tgt_shares = ['AARTIIND', 'ABB', 'ABCAPITAL', 'ABFRL', 'ADANIENT', 'ADANIPORTS', 'ALKEM', 'AMARAJABAT', 'AMBUJACEM',
              'APLLTD', 'APOLLOHOSP', 'APOLLOTYRE', 'ASHOKLEY', 'ASTRAL', 'ATUL', 'AUBANK', 'AUROPHARMA', 'BAJAJFINSV',
              'BAJFINANCE', 'BALKRISIND', 'BALRAMCHIN', 'BANDHANBNK', 'BANKBARODA', 'BATAINDIA', 'BEL', 'BHARATFORG',
              'BIOCON', 'BRITANNIA', 'BSOFT', 'CANBK', 'CANFINHOME', 'CHAMBLFERT', 'CHOLAFIN', 'CIPLA', 'COFORGE',
              'CONCOR', 'COROMANDEL', 'CROMPTON', 'CUMMINSIND', 'DABUR', 'DALBHARAT', 'DEEPAKNTR', 'DELTACORP',
              'DIVISLAB', 'DIXON', 'DLF', 'DRREDDY', 'ESCORTS', 'EXIDEIND', 'GLENMARK', 'GLS', 'GNFC', 'GODREJCP',
              'GODREJPROP', 'GRANULES', 'GRASIM', 'GUJGASLTD', 'HAL', 'HAVELLS', 'HCLTECH', 'HDFCAMC', 'HDFCLIFE',
              'HINDALCO', 'HINDCOPPER', 'ICICIGI', 'ICICIPRULI', 'IEX', 'IGL', 'INDHOTEL', 'INDIACEM', 'INDIAMART',
              'INDIGO', 'INDUSINDBK', 'INDUSTOWER', 'INTELLECT', 'IPCALAB', 'JINDALSTEL', 'JKCEMENT', 'JSWSTEEL',
              'JUBLFOOD', 'KOTAKBANK', 'LALPATHLAB', 'LAURUSLABS', 'LICHSGFIN', 'LTIM', 'LTTS', 'LUPIN',
              'M&MFIN', 'MANAPPURAM', 'MARICO', 'MCDOWELL-N', 'MCX', 'METROPOLIS', 'MFSL', 'MGL', 'MPHASIS',
              'MUTHOOTFIN', 'NAM-INDIA', 'NAUKRI', 'NAVINFLUOR', 'NMDC', 'NTPC', 'OBEROIRLTY', 'PEL', 'PERSISTENT',
              'PETRONET', 'PIDILITIND', 'POLYCAB', 'POWERGRID', 'RAIN', 'RAMCOCEM', 'RBLBANK', 'RECLTD',
              'SBICARD', 'SBILIFE', 'SIEMENS', 'SRF', 'STAR', 'SUNPHARMA', 'SYNGENE', 'TATACOMM', 'TECHM',
              'TORNTPHARM', 'TORNTPOWER', 'TRENT', 'TVSMOTOR', 'UBL', 'ULTRACEMCO', 'UPL', 'VEDL', 'VOLTAS', 'ZEEL',
              'ZYDUSLIFE']

# creating output workbook to save the data
op_wb = xl.Workbook()

del op_wb['Sheet']          # deleting default sheet

all_tr_sheet = op_wb.create_sheet("All Trades")
actual_tr_sheet = op_wb.create_sheet("Actual Trades")

row = 2     # output workbook starting row

df = pd.read_csv(md_path_zipped[:-4])

# Fixed headings to be put in all_tr_sheet sheet
all_tr_sheet.cell(1, 1).value = "Share Name"
all_tr_sheet.cell(1, 2).value = "High Value"
all_tr_sheet.cell(1, 3).value = "Low Value"

all_tr_sheet.cell(1, 5).value = "Buy Entry"
all_tr_sheet.cell(1, 6).value = "Buy Target"
all_tr_sheet.cell(1, 7).value = "Buy Stoploss"
all_tr_sheet.cell(1, 8).value = "Buy Quantity"

all_tr_sheet.cell(1, 10).value = "Sell Entry"
all_tr_sheet.cell(1, 11).value = "Sell Target"
all_tr_sheet.cell(1, 12).value = "Sell Stoploss"
all_tr_sheet.cell(1, 13).value = "Sell Quantity"
all_tr_sheet.cell(1, 15).value = "3:30 Price"


# Fixed headings to be put in actual_tr_sheet sheet
actual_tr_sheet.cell(1, 1).value = "Share Name"
actual_tr_sheet.cell(1, 2).value = "New High Value"
actual_tr_sheet.cell(1, 3).value = "New Low Value"

actual_tr_sheet.cell(1, 5).value = "Entry Price"
actual_tr_sheet.cell(1, 6).value = "Trade Type"
actual_tr_sheet.cell(1, 7).value = "Target Price"
actual_tr_sheet.cell(1, 8).value = "Stoploss Price"
actual_tr_sheet.cell(1, 9).value = "3:30 Price"
actual_tr_sheet.cell(1, 10).value = "Exit P&L"

actual_tr_sheet.cell(1, 12).value = "Peak Exit P&L"

actual_tr_sheet.cell(1, 14).value = "3:30 Exit P&L"

for i in range(len(df)):
    share = str(df.iloc[i, 0])          # share name
    high = df.iloc[i, 3]                # high value
    low = df.iloc[i, 4]                 # low value
    symbol = df.iloc[i, 1]              # symbol value
    val_330 = df.iloc[i, 6]             # last value

    if share in tgt_shares and symbol == 'EQ':
        all_tr_sheet.cell(row, 1).value = share
        all_tr_sheet.cell(row, 2).value = high
        all_tr_sheet.cell(row, 3).value = low

        # adding trade entry prices and quantities
        b_entry = round_up((high + (high * 0.01)))
        b_tgt = round_down((b_entry + (b_entry * 0.015)))
        b_sl = round_up((b_entry - (b_entry * 0.0125)))
        b_qty = 100000 // b_entry

        s_entry = round_down((low - (low * 0.01)))
        s_tgt = round_up((s_entry - (s_entry * 0.015)))
        s_sl = round_down((s_entry + (s_entry * 0.0125)))
        s_qty = 100000 // s_entry

        # BUY
        all_tr_sheet.cell(row, 5).value = b_entry           # Buy Entry
        all_tr_sheet.cell(row, 6).value = b_tgt             # Buy Target
        all_tr_sheet.cell(row, 7).value = b_sl              # Buy Stoploss
        all_tr_sheet.cell(row, 8).value = b_qty             # Buy Quantity

        # SELL
        all_tr_sheet.cell(row, 10).value = s_entry           # Sell Entry
        all_tr_sheet.cell(row, 11).value = s_tgt             # Sell Target
        all_tr_sheet.cell(row, 12).value = s_sl              # Sell Stoploss
        all_tr_sheet.cell(row, 13).value = s_qty             # Sell Quantity

        # 3:30 Value
        all_tr_sheet.cell(row, 15).value = val_330           # LTP/ 3:30 close value

        row += 1


# # formatting file
# for i in range(200):
#     for j in range(20):
#         all_tr_sheet.cell(i + 1, j + 1).number_format = '0.00'

op_wb.save(rf"E:\Daily Data work\Trading Algorithm\{yr}\{mnth}\{date[:2]}{mnth}20{date[6:]}algo.xlsx")        # saving our file in Daily Data work
#
# # auto fitting columns
# excel = Dispatch('Excel.Application')
# wb = excel.Workbooks.Open(rf"E:\Daily Data work\Trading Algorithm\{yr}\{mnth}\{date[:2]}{mnth}20{date[6:]}algo.xlsx")
#
# # Activating sheets
# excel.Worksheets(1).Activate()
# excel.ActiveSheet.Columns.AutoFit()
# excel.Worksheets(2).Activate()
# excel.ActiveSheet.Columns.AutoFit()
#
# # saving
# wb.Save()
# wb.Close()
#
# # loading yesterday's workbook
# old_wb = xl.load_workbook(rf"E:\Daily Data work\Trading Algorithm\{old_yr}\{old_mnth}\{old_date[:2]}{old_mnth}20{old_date[6:]}algo.xlsx")
# old_all_tr_sheet = old_wb["All Trades"]
# old_actual_tr_sheet = old_wb["Actual Trades"]
#
# # loading current day's algorithm workbook and sheets
# today_wb = xl.load_workbook(rf"E:\Daily Data work\Trading Algorithm\{yr}\{mnth}\{date[:2]}{mnth}20{date[6:]}algo.xlsx")
# new_all_tr_sheet = today_wb["All Trades"]
#
#
# old_all_row = 2                     # row for previous day's workbook's 'All Trades' sheet
# old_actual_row = 2              # row for previous day's workbook's 'Actual Trades' sheet
#
# trade_type = ''
# exit_type = ''          # exit on target, stoploss or 3:30 pm
#
# while old_all_row < 135:       # WILL NEED TO MAKE IT DYNAMIC OR CHANGE WHEN ADDING NEW SHARES
#     old_buy_entry = old_all_tr_sheet.cell(old_all_row, 5).value             # old buy entry
#     old_sell_entry = old_all_tr_sheet.cell(old_all_row, 10).value           # old sell entry
#     new_330_close = new_all_tr_sheet.cell(old_all_row, 15).value            # 3:30 close value of today
#
#     buy_tgt = old_all_tr_sheet.cell(old_all_row, 6).value                   # buy target
#     buy_sl = old_all_tr_sheet.cell(old_all_row, 7).value                    # buy stoploss
#     sell_tgt = old_all_tr_sheet.cell(old_all_row, 11).value                 # sell target
#     sell_sl = old_all_tr_sheet.cell(old_all_row, 12).value                  # sell stoploss
#     buy_qty = old_all_tr_sheet.cell(old_all_row, 8).value                   # buy quantity
#     sell_qty = old_all_tr_sheet.cell(old_all_row, 13).value                 # sell quantity
#
#     new_high = new_all_tr_sheet.cell(old_all_row, 2).value                  # new high
#     new_low = new_all_tr_sheet.cell(old_all_row, 3).value                   # new low
#
#     # checking trade entry conditions
#     if new_low <= old_sell_entry:
#         trade_type = 'S'        # entering sell trade
#     elif new_high >= old_buy_entry:
#         trade_type = 'B'        # entering buy trade
#     else:
#         trade_type = 'N'
#
#     if trade_type == 'N':           # no trade entered
#         old_all_row += 1
#         continue
#
#     elif trade_type == 'B':         # buy entered
#         if new_high >= buy_tgt:
#             exit_type = 'tgt'
#         elif buy_sl >= new_low:             # between high and low
#             exit_type = 'sl'
#         else:
#             exit_type = '330'
#
#     elif trade_type == 'S':         # sell entered
#         if new_low <= sell_tgt:
#             exit_type = 'tgt'
#         elif sell_sl <= new_high:           # between high and low
#             exit_type = 'sl'
#         else:
#             exit_type = '330'
#
#     else:
#         print("You shouldn't be here")          # if code reaches here, there is some logic error
#
#     # copying values
#     old_actual_tr_sheet.cell(old_actual_row, 1).value = old_all_tr_sheet.cell(old_all_row, 1).value     # name
#     old_actual_tr_sheet.cell(old_actual_row, 2).value = new_all_tr_sheet.cell(old_all_row, 2).value     # next day's high
#     old_actual_tr_sheet.cell(old_actual_row, 3).value = new_all_tr_sheet.cell(old_all_row, 3).value     # next day's low
#
#     old_actual_tr_sheet.cell(old_actual_row, 6).value = "BUY" if trade_type == 'B' else "SELL"          # BUY/SELL
#     old_actual_tr_sheet.cell(old_actual_row, 5).value = old_buy_entry if trade_type == 'B' else old_sell_entry          # BUY/SELL Price
#
#     net = 0         # net profit/loss
#
#     if trade_type == 'B':       # if trade is a buy trade
#         if exit_type == 'tgt':
#             old_actual_tr_sheet.cell(old_actual_row, 7).value = buy_tgt
#             net = (buy_tgt - old_buy_entry) * buy_qty           # profit points * qty
#         if exit_type == 'sl':
#             old_actual_tr_sheet.cell(old_actual_row, 8).value = buy_sl
#             net = (buy_sl - old_buy_entry) * buy_qty            # loss points * qty
#
#         old_actual_tr_sheet.cell(old_actual_row, 10).value = net  # putting net P/L values in sheet
#
#         # if we had taken max profit
#         net = (new_high - old_buy_entry) * buy_qty
#         old_actual_tr_sheet.cell(old_actual_row, 12).value = net
#
#         # P/L if exit on 3:30 pm
#         net = (new_330_close - old_buy_entry) * buy_qty
#         old_actual_tr_sheet.cell(old_actual_row, 14).value = net
#
#         old_actual_row += 1
#
#     if trade_type == 'S':       # if trade is a sell trade
#         if exit_type == 'tgt':
#             old_actual_tr_sheet.cell(old_actual_row, 7).value = sell_tgt
#             net = (old_sell_entry - sell_tgt) * sell_qty        # profit points * qty
#         if exit_type == 'sl':
#             old_actual_tr_sheet.cell(old_actual_row, 8).value = sell_sl
#             net = (old_sell_entry - sell_sl) * sell_qty         # loss points * qty
#
#         old_actual_tr_sheet.cell(old_actual_row, 10).value = net       # putting net P/L values in sheet
#
#         # if we had taken max profit
#         net = (old_sell_entry - new_low) * sell_qty
#         old_actual_tr_sheet.cell(old_actual_row, 12).value = net
#
#         # P/L if exit on 3:30 pm
#         net = (old_sell_entry - new_330_close) * sell_qty
#         old_actual_tr_sheet.cell(old_actual_row, 14).value = net
#
#         old_actual_row += 1
#
#     old_all_row += 1
#
# # finally saving the old workbook
# old_wb.save(rf"E:\Daily Data work\Trading Algorithm\{old_yr}\{old_mnth}\{old_date[:2]}{old_mnth}20{old_date[6:]}algo.xlsx")
#
# # saving current date in file
# # with open('date.txt', 'w') as file:
# #     file.write(f"{date}\n{mnth}\n{yr}")
#
#
# # auto fitting excel files
# excel = Dispatch('Excel.Application')
# wb = excel.Workbooks.Open(rf"E:\Daily Data work\Trading Algorithm\{old_yr}\{old_mnth}\{old_date[:2]}{old_mnth}20{old_date[6:]}algo.xlsx")
#
# # Activating sheets
# excel.Worksheets(1).Activate()
# excel.ActiveSheet.Columns.AutoFit()
# excel.Worksheets(2).Activate()
# excel.ActiveSheet.Columns.AutoFit()
#
# # saving
# wb.Save()
# wb.Close()



