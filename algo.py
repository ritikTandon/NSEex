import shutil

import openpyxl as xl
import os
import requests
from openpyxl.styles import PatternFill
import datetime

from xls2xlsx import XLS2XLSX

from date_variables import date, mnth, yr
from time import sleep
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException

# headers = {'User-Agent': 'Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:52.0) Gecko/20100101 Firefox/52.0'}

cash_share_list = ['AARTIIND', 'ABB', 'ABCAPITAL', 'ABFRL', 'ADANIENT', 'ADANIPORTS', 'ALKEM', 'AMBUJACEM',
                     'APOLLOHOSP', 'APOLLOTYRE', 'ASHOKLEY', 'ASTRAL', 'ATUL', 'AUBANK', 'AUROPHARMA', 'BAJAJFINSV',
                     'BAJFINANCE', 'BALKRISIND', 'BALRAMCHIN', 'BANDHANBNK', 'BANKBARODA', 'BATAINDIA', 'BEL',
                     'BHARATFORG', 'BIOCON', 'BRITANNIA', 'BSOFT', 'CANBK', 'CANFINHOME', 'CHAMBLFERT', 'CHOLAFIN',
                     'CIPLA', 'COFORGE', 'CONCOR', 'COROMANDEL', 'CROMPTON', 'CUMMINSIND', 'DABUR', 'DALBHARAT',
                     'DEEPAKFERT', 'DEEPAKNTR', 'DELTACORP', 'DIVISLAB', 'DIXON', 'DLF', 'DRREDDY', 'ESCORTS',
                     'EXIDEIND', 'GLENMARK', 'GLS', 'GNFC', 'GODREJCP', 'GODREJPROP', 'GRANULES', 'GRASIM', 'GUJGASLTD',
                     'HAL', 'HAVELLS', 'HCLTECH', 'HDFCAMC', 'HDFCLIFE', 'HINDALCO', 'HINDCOPPER', 'ICICIGI',
                     'ICICIPRULI', 'IEX', 'IGL', 'INDHOTEL', 'INDIACEM', 'INDIAMART', 'INDIGO', 'INDUSINDBK',
                     'INDUSTOWER', 'INTELLECT', 'IPCALAB', 'JINDALSTEL', 'JKCEMENT', 'JSWSTEEL', 'JUBLFOOD',
                     'KOTAKBANK', 'LALPATHLAB', 'LAURUSLABS', 'LICHSGFIN', 'LTIM', 'LTTS', 'LUPIN', 'M&MFIN',
                     'MANAPPURAM', 'MARICO', 'MCDOWELL-N', 'MCX', 'METROPOLIS', 'MFSL', 'MGL', 'MPHASIS', 'MUTHOOTFIN',
                     'NAM-INDIA', 'NAUKRI', 'NAVINFLUOR', 'NMDC', 'NTPC', 'OBEROIRLTY', 'PEL', 'PERSISTENT', 'PETRONET',
                     'PIDILITIND', 'POLYCAB', 'POWERGRID', 'RAIN', 'RAMCOCEM', 'RBLBANK', 'RECLTD', 'SBICARD',
                     'SBILIFE', 'SIEMENS', 'SRF', 'STAR', 'SUNPHARMA', 'SYNGENE', 'TATACOMM', 'TATAMOTORS', 'TECHM',
                     'TORNTPHARM', 'TORNTPOWER', 'TRENT', 'TVSMOTOR', 'UBL', 'ULTRACEMCO', 'UPL', 'VEDL', 'VOLTAS',
                     'ZEEL', 'ZYDUSLIFE']

# cash_share_list = ["ADANI"]

cashHL_wb = xl.load_workbook(r'C:\Users\admin\PycharmProjects\daily data\algo high low.xlsx')
cashHL_sheet = cashHL_wb['Sheet1']
cashHL_row = 2


# converting xls to xlsx for algo sheet
x2x = XLS2XLSX(r'E:\Daily Data work\algo.xls')

wb = x2x.to_xlsx()
wb.save(r'E:\Daily Data work\algo.xlsx')

csh_wb = xl.load_workbook(r'E:\Daily Data work\algo.xlsx')
csh_sheet = csh_wb['algo-Sheet1']
csh_row = 2

fl_9_25 = 0.392361111       # 9:25 time value in general format

# copying hourlys (.xls) as backup
# path to source directory
src_dir = rf"E:\Daily Data work\hourlys 1 minute ALGO\{yr}\{mnth}\{date}"

# path to destination directory
dest_dir = rf"C:\Users\admin\PycharmProjects\daily data\Daily Backup hourlys\1 min algo"

# getting all the files in the source directory
src_files = os.listdir(src_dir)
for file_name in src_files:
    full_file_name = os.path.join(src_dir, file_name)
    if os.path.isfile(full_file_name):
        shutil.copy(full_file_name, dest_dir)
print("Files copied as backup!")


for share in cash_share_list:
    path = rf"E:\Daily Data work\hourlys 1 minute ALGO\{yr}\{mnth}\{date}\{share}.xlsx"
    xls_path = rf"E:\Daily Data work\hourlys 1 minute ALGO\{yr}\{mnth}\{date}\{share}.xls"
    # try:      # cant use this skip unless I update cash high low values of a share and reload wb after every loop
    #     x2x = XLS2XLSX(xls_path)
    # except FileNotFoundError:
    #     continue
    x2x = XLS2XLSX(xls_path)

    wb = x2x.to_xlsx()
    # wb = xl.load_workbook(path)
    sheet = wb[f"{share}-Sheet1"]

    start_row = 2
    time_cell = sheet.cell(start_row, 7)

    while time_cell.value < fl_9_25:
        start_row += 1
        time_cell = sheet.cell(start_row, 7)

    print(f"starting row is {start_row}")

    # saving starting row value so I can read this value later on when making algo data
    sheet.cell(1, 18).value = start_row

    start_row_2 = start_row
    start_row = 2  # converting time from start

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
    while cur_time and cur_time <= end_time:
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

    # high
    cashHL_sheet.cell(cashHL_row, 2).value = HIGH

    # low
    cashHL_sheet.cell(cashHL_row, 3).value = LOW

    # LTP
    cashHL_sheet.cell(cashHL_row, 5).value = csh_sheet.cell(csh_row, 2).value

    # vol
    volume = csh_sheet.cell(csh_row, 5).value
    volume = volume // 100000

    cashHL_sheet.cell(cashHL_row, 6).value = volume

    # 9:25 close
    cashHL_sheet.cell(cashHL_row, 7).value = cl_9_25

    cashHL_row += 1
    csh_row += 1

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
            # if HIGH == 375:
            #     print('k')

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

    sheet.freeze_panes = sheet["A2"]

    wb.save(path)
    os.remove(xls_path)

    print(f"{share} done")

# for close filling
options = Options()
# options.add_argument('--headless=new')
options.add_argument("--disable-blink-features=AutomationControlled")

# Exclude the collection of enable-automation switches
options.add_experimental_option("excludeSwitches", ["enable-automation"])

# Turn-off userAutomationExtension
options.add_experimental_option("useAutomationExtension", False)

cash_close_list = ['AARTIIND', 'ABB', 'ABCAPITAL', 'ABFRL', 'ADANIENT', 'ADANIPORTS', 'ALKEM', 'AMBUJACEM',
                     'APOLLOHOSP', 'APOLLOTYRE', 'ASHOKLEY', 'ASTRAL', 'ATUL', 'AUBANK', 'AUROPHARMA', 'BAJAJFINSV',
                     'BAJFINANCE', 'BALKRISIND', 'BALRAMCHIN', 'BANDHANBNK', 'BANKBARODA', 'BATAINDIA', 'BEL',
                     'BHARATFORG', 'BIOCON', 'BRITANNIA', 'BSOFT', 'CANBK', 'CANFINHOME', 'CHAMBLFERT', 'CHOLAFIN',
                     'CIPLA', 'COFORGE', 'CONCOR', 'COROMANDEL', 'CROMPTON', 'CUMMINSIND', 'DABUR', 'DALBHARAT',
                     'DEEPAKFERT', 'DEEPAKNTR', 'DELTACORP', 'DIVISLAB', 'DIXON', 'DLF', 'DRREDDY', 'ESCORTS',
                     'EXIDEIND', 'GLENMARK', 'GLS', 'GNFC', 'GODREJCP', 'GODREJPROP', 'GRANULES', 'GRASIM', 'GUJGASLTD',
                     'HAL', 'HAVELLS', 'HCLTECH', 'HDFCAMC', 'HDFCLIFE', 'HINDALCO', 'HINDCOPPER', 'ICICIGI',
                     'ICICIPRULI', 'IEX', 'IGL', 'INDHOTEL', 'INDIACEM', 'INDIAMART', 'INDIGO', 'INDUSINDBK',
                     'INDUSTOWER', 'INTELLECT', 'IPCALAB', 'JINDALSTEL', 'JKCEMENT', 'JSWSTEEL', 'JUBLFOOD',
                     'KOTAKBANK', 'LALPATHLAB', 'LAURUSLABS', 'LICHSGFIN', 'LTIM', 'LTTS', 'LUPIN', 'M%26MFIN',
                     'MANAPPURAM', 'MARICO', 'UNITDSPR', 'MCX', 'METROPOLIS', 'MFSL', 'MGL', 'MPHASIS', 'MUTHOOTFIN',
                     'NAM-INDIA', 'NAUKRI', 'NAVINFLUOR', 'NMDC', 'NTPC', 'OBEROIRLTY', 'PEL', 'PERSISTENT', 'PETRONET',
                     'PIDILITIND', 'POLYCAB', 'POWERGRID', 'RAIN', 'RAMCOCEM', 'RBLBANK', 'RECLTD', 'SBICARD',
                     'SBILIFE', 'SIEMENS', 'SRF', 'STAR', 'SUNPHARMA', 'SYNGENE', 'TATACOMM', 'TATAMOTORS', 'TECHM',
                     'TORNTPHARM', 'TORNTPOWER', 'TRENT', 'TVSMOTOR', 'UBL', 'ULTRACEMCO', 'UPL', 'VEDL', 'VOLTAS',
                     'ZEEL', 'ZYDUSLIFE']

# cash_close_list1 = ["M%26M"]

manual = []         # list to keep track of the shares whose values selenium couldn't get
close = []
# close = ['2446.95', '398.5', '1478.0', '7005.0', '227.65', '188.4', '234.15', '470.1', '3337.8', '132.8', '1170.0',
# '1608.5', '448.85', '959.0', '1387.8', '1392.95', '662.9', '421.45', '1544.95', '1544.95', '2573.2', '561.0', '548.9',
# '1000.05', '606.8', '230.95', '117.85', '8024.95']

for share in cash_close_list:
    driver = webdriver.Chrome(options=options)

    driver.get(f"https://www.nseindia.com/get-quotes/equity?symbol={share}")

    try:
        sleep(2)
        myElem = WebDriverWait(driver, 20).until(ec.presence_of_element_located((By.ID, 'quoteLtp')))
        # sleep(5)
        close_val = driver.find_element(By.ID, "quoteLtp").text

        while close_val == '':
            driver.refresh()
            WebDriverWait(driver, 10).until(ec.presence_of_element_located((By.ID, 'quoteLtp')))
            close_val = driver.find_element(By.ID, "quoteLtp").text
            sleep(0.5)

        close_val = close_val.replace(",", "")

        # truncating last 0
        if close_val[len(close_val)-1:len(close_val)] == '0':
            close_val = close_val[:len(close_val)-1]

        close.append(close_val)

        print(f'{share}: {close_val}')
        if close_val == '':
            manual.append(share)

    except TimeoutException:
        print(f"Loading took too much time for {share}!")
        close.append('')
        manual.append(share)

    driver.close()

print(close)
print(manual)

i = 0
# close = ['627.0', '8318.0', '222.0', '263.95', '3034.05', '1342.0', '5301.15', '615.0', '5920.0', '489.0', '206.55', '2276.0', '5900.1', '621.5', '1160.1', '1593.95', '6731.0', '2590.0', '378.4', '180.2', '263.8', '1334.95', '239.75', '1489.75', '308.95', '5142.8', '616.9', '113.9', '762.0', '403.5', '1268.8', '1420.65', '4675.9', '1037.9', '1246.7', '338.6', '3737.3', '543.0', '1793.45', '557.9', '2453.0', '120.9', '3918.3', '8235.0', '845.05', '5829.0', '3749.0', '473.25', '1009.0', '825.6', '666.0', '1301.05', '2872.0', '402.0', '2379.0', '553.15', '4637.9', '1776.9', '1347.15', '3855.9', '566.1', '652.3', '378.5', '1675.05', '590.15', '148.4', '441.5', '568.25', '208.6', '2633.0', '4299.9', '1410.0', '341.5', '879.9', '1286.35', '1005.5', '3864.0', '884.0', '471.5', '1670.9', '2515.9', '445.0', '655.8', '4764.1', '4515.0', '1658.55', '267.75', '185.7', '593.9', '1179.0', '3940.0', '1862.0', '1000.95', '1312.9', '2370.0', '1694.0', '582.0', '5875.3', '3310.65', '265.1', '361.0', '1721.0', '826.4', '3518.0', '312.5', '3017.35', '6478.0', '312.3', '161.9', '768.75', '252.05', '541.1', '713.9', '1450.2', '7070.0', '2274.0', '869.5', '1530.05', '687.1', '1800.0', '939.2', '1304.0', '2703.95', '1345.9', '4596.5', '2124.3', '1902.5', '9700.0', '510.15', '433.6', '1322.95', '133.0', '1009.0']
while i < len(cash_close_list):
    close_cell = cashHL_sheet.cell(i+2, 4)

    if close[i] == '':
        close_cell.value = 0

    else:
        close_cell.value = float(close[i])
        close_cell.number_format = "0.00"

    i += 1

cashHL_wb.save(r'C:\Users\admin\PycharmProjects\daily data\algo high low.xlsx')
