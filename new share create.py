import shutil
import pandas as pd
import glob
import os
import pyautogui as pg
import pyperclip as pc
from date_variables import date, mnth, yr
from time import sleep
import time
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

# for close filling
options = Options()
# options.add_argument('--headless=new')
options.add_argument("--disable-blink-features=AutomationControlled")

# Exclude the collection of enable-automation switches
options.add_experimental_option("excludeSwitches", ["enable-automation"])

# maximizing window
options.add_argument("--start-maximized")

# Turn-off userAutomationExtension
options.add_experimental_option("useAutomationExtension", False)

# share_add_list = ['AARTIIND', 'ABB', 'ABCAPITAL', 'ABFRL', 'ADANIENT', 'ADANIPORTS', 'ALKEM', 'AMBUJACEM',
#                      'APOLLOHOSP', 'APOLLOTYRE', 'ASHOKLEY', 'ASTRAL', 'ATUL', 'AUBANK', 'AUROPHARMA', 'BAJAJFINSV',
#                      'BAJFINANCE', 'BALKRISIND', 'BALRAMCHIN', 'BANDHANBNK', 'BANKBARODA', 'BATAINDIA', 'BEL',
#                      'BHARATFORG', 'BIOCON', 'BRITANNIA', 'BSOFT', 'CANBK', 'CANFINHOME', 'CHAMBLFERT', 'CHOLAFIN',
#                      'CIPLA', 'COFORGE', 'CONCOR', 'COROMANDEL', 'CROMPTON', 'CUMMINSIND', 'DABUR', 'DALBHARAT',
#                      'DEEPAKFERT', 'DEEPAKNTR', 'DELTACORP', 'DIVISLAB', 'DIXON', 'DLF', 'DRREDDY', 'ESCORTS',
#                      'EXIDEIND', 'GLENMARK', 'GLS', 'GNFC', 'GODREJCP', 'GODREJPROP', 'GRANULES', 'GRASIM', 'GUJGASLTD',
#                      'HAL', 'HAVELLS', 'HCLTECH', 'HDFCAMC', 'HDFCLIFE', 'HINDALCO', 'HINDCOPPER', 'ICICIGI',
#                      'ICICIPRULI', 'IEX', 'IGL', 'INDHOTEL', 'INDIACEM', 'INDIAMART', 'INDIGO', 'INDUSINDBK',
#                      'INDUSTOWER', 'INTELLECT', 'IPCALAB', 'JINDALSTEL', 'JKCEMENT', 'JSWSTEEL', 'JUBLFOOD',
#                      'KOTAKBANK', 'LALPATHLAB', 'LAURUSLABS', 'LICHSGFIN', 'LTIM', 'LTTS', 'LUPIN', 'M&MFIN',
#                      'MANAPPURAM', 'MARICO', 'MCDOWELL-N', 'MCX', 'METROPOLIS', 'MFSL', 'MGL', 'MPHASIS', 'MUTHOOTFIN',
#                      'NAM-INDIA', 'NAUKRI', 'NAVINFLUOR', 'NMDC', 'NTPC', 'OBEROIRLTY', 'PEL', 'PERSISTENT', 'PETRONET',
#                      'PIDILITIND', 'POLYCAB', 'POWERGRID', 'RAIN', 'RAMCOCEM', 'RBLBANK', 'RECLTD', 'SBICARD',
#                      'SBILIFE', 'SIEMENS', 'SRF', 'STAR', 'SUNPHARMA', 'SYNGENE', 'TATACOMM', 'TATAMOTORS', 'TECHM',
#                      'TORNTPHARM', 'TORNTPOWER', 'TRENT', 'TVSMOTOR', 'UBL', 'ULTRACEMCO', 'UPL', 'VEDL', 'VOLTAS',
#                      'ZEEL', 'ZYDUSLIFE']

share_add_list = ['HINDALCO', 'NTPC']
# share_add_list = ['HINDALCO']

coords = {"date left": (1048, 512), "date right": (1348, 512), "left year": (982, 555), "right year": (1273, 555),
          "left left": (887, 552), "left jan": (898, 596), "right jan": (1184, 596), "jan 1 2020": (984, 614),
          "right left": (1172, 555), "nov 2023 right": (1294, 707), "left right": (1079, 555),
          "right right": (1368, 555),
          "jan 1 2021 right": (1331, 617), "jan 1 2021 left": (1048, 614), "jan 1 2022 right": (1366, 616),
          "jan 1 2022 left": (1081, 618), "jan 1 2023 right": (1174, 613), "jan 1 2023 left": (890, 617),
          "13 nov 2023 right": (1205, 680), "filter": (1460, 508), "download csv": (1467, 565)}


for share in share_add_list:
    driver = webdriver.Chrome(options=options)

    driver.get(f"https://www.nseindia.com/get-quotes/equity?symbol={share}")

    try:
        sleep(2)
        myElem = WebDriverWait(driver, 10).until(ec.presence_of_element_located((By.ID, 'historic_data')))
        # sleep(5)
        driver.find_element(By.ID, 'historic_data').click()

        sleep(2)
        pg.click(coords["date left"])
        sleep(2)
        pg.click(coords["left year"])
        sleep(2)
        pg.click(coords["left left"], clicks=3, interval=0.5)
        sleep(2)
        pg.click(coords["left jan"])
        sleep(2)
        pg.click(coords["jan 1 2020"])
        sleep(2)

        pg.click(coords["date right"])
        sleep(2)
        pg.click(coords["right year"])
        sleep(2)
        pg.click(coords["right left"], clicks=2, interval=0.5)
        sleep(2)
        pg.click(coords["right jan"])
        sleep(2)
        pg.click(coords["jan 1 2021 right"])
        sleep(2)

        pg.click(coords["filter"])
        sleep(2)
        pg.click(coords["download csv"])
        sleep(2)

        # 2nd csv
        pg.click(coords["date left"])
        sleep(2)
        pg.click(coords["left year"])
        sleep(2)
        pg.click(coords["left right"])
        sleep(2)
        pg.click(coords["left jan"])
        sleep(2)
        pg.click(coords["jan 1 2021 left"])
        sleep(2)

        pg.click(coords["date right"])
        sleep(2)
        pg.click(coords["right year"])
        sleep(2)
        pg.click(coords["right right"])
        sleep(2)
        pg.click(coords["right jan"])
        sleep(2)
        pg.click(coords["jan 1 2022 right"])
        sleep(2)
        pg.click(coords["filter"])
        sleep(2)
        pg.click(coords["download csv"])
        sleep(2)

        # jan 1 2022 - jan 1 2023
        pg.click(coords["date left"])
        sleep(2)
        pg.click(coords["left year"])
        sleep(2)
        pg.click(coords["left right"])
        sleep(2)
        pg.click(coords["left jan"])
        sleep(2)
        pg.click(coords["jan 1 2022 left"])
        sleep(2)

        pg.click(coords["date right"])
        sleep(2)
        pg.click(coords["right year"])
        sleep(2)
        pg.click(coords["right right"])
        sleep(2)
        pg.click(coords["right jan"])
        sleep(2)
        pg.click(coords["jan 1 2023 right"])
        sleep(2)
        pg.click(coords["filter"])
        sleep(2)
        pg.click(coords["download csv"])
        sleep(2)

        # jan 1 2023 - nov 13 2023
        pg.click(coords["date left"])
        sleep(2)
        pg.click(coords["left year"])
        sleep(2)
        pg.click(coords["left right"])
        sleep(2)
        pg.click(coords["left jan"])
        sleep(2)
        pg.click(coords["jan 1 2023 left"])
        sleep(2)

        pg.click(coords["date right"])
        sleep(2)
        pg.click(coords["right year"])
        sleep(2)
        # pg.click(coords["right right"])
        # sleep(2)
        pg.click(coords["nov 2023 right"])
        sleep(2)
        pg.click(coords["13 nov 2023 right"])
        sleep(2)
        pg.click(coords["filter"])
        sleep(2)
        pg.click(coords["download csv"])
        sleep(2)

    except TimeoutException:
        print("Loading took too much time!")

    # saving files as one
    # merging the files
    joined_files = os.path.join("C:\\Users\\admin\\Downloads", "Quote-Equity*.csv")

    # A list of all joined files is returned
    joined_list = glob.glob(joined_files)

    # Finally, the files are joined
    df = pd.concat(map(pd.read_csv, joined_list), ignore_index=True)

    # converting date column to datetime and sorting via that column
    df["Date "] = pd.to_datetime(df["Date "])
    df.sort_values('Date ', axis=0, inplace=True)

    # converting volume to number and then to lakhs
    df.replace(',', '', regex=True, inplace=True)
    df['VOLUME '] = df['VOLUME '].apply(pd.to_numeric, errors='coerce')
    df['VOLUME '] = df['VOLUME '] // 100000

    # renaming columns
    df.rename(
        columns={'Date ': 'Date', 'ltp ': 'Close', 'close ': 'LTP', 'HIGH ': 'High', 'LOW ': 'Low', 'VOLUME ': 'VOL'},
        inplace=True)

    # dropping duplicate date rows and diwali row
    df.drop_duplicates(subset="Date", keep='first', inplace=True)
    df.drop(df.index[-2], inplace=True)
    df.drop(df.index[220], inplace=True)  # 14 nov 2020

    # only keeping necessary columns
    df = df.iloc[:, [0, 3, 4, 6, 7, 11]]
    # print(df.to_string(max_cols=11))

    # saving to excel
    writer = pd.ExcelWriter(rf"C:\Users\admin\Downloads\current\{share}.xlsx", engine="xlsxwriter",
                            datetime_format="dd-mmm-yy")

    # Convert the dataframe to an XlsxWriter Excel object.
    df.to_excel(writer, sheet_name="D", index=False)
    wb1 = writer.book
    writer.close()

    insert_rows = [40, 52, 69, 71, 75, 77, 90, 106, 200, 231, 241, 260, 282, 314, 326, 330, 338, 343, 359, 408, 429,
                   445, 470, 485, 495, 543, 569, 580, 599, 600, 612, 682, 686, 698, 723, 738, 747, 804, 832, 849, 852,
                   855, 860, 871, 914, 947, 972, 981, 997]

    wb = xl.load_workbook(rf'C:\Users\admin\Downloads\current\{share}.xlsx')
    sheet = wb['D']

    for row in insert_rows:
        sheet.insert_rows(row)

    # @TODO CHANGE THIS FOR ALGO SHARES!!!!!!!!!!!!!!!!!!
    wb.save(rf'E:\Daily Data work\CASH\{share}.xlsx')

    # deleting .csv files from current after making a consolidated file
    dir_name = r'C:\Users\admin\Downloads'
    test = os.listdir(dir_name)

    for item in test:
        if item.endswith(".csv"):
            os.remove(os.path.join(dir_name, item))
