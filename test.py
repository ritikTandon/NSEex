# from time import sleep
# from selenium import webdriver
# from selenium.webdriver.support.ui import WebDriverWait
# from selenium.webdriver.support import expected_conditions as ec
# from selenium.webdriver.common.by import By
# from selenium.webdriver.chrome.options import Options
# from selenium.common.exceptions import TimeoutException
#
# # headers = {'User-Agent': 'Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:52.0) Gecko/20100101 Firefox/52.0'}
#
# cash_share_list = ["AARTIIND", "ADANI", "APOLLO", "BAJFINSV", "BAJFIN", "BANBK", "BARODA", "COALIND", "DLF", "EICHER",
#                    "FEDBANK", "HCL", "HDFC", "HIND", "ICICI", "INDUSIND", "INFY", "JIND", "LIC", "M&M", "M&MFIN", "NTPC",
#                    "REL", "SBIN", "SUNTV", "TCHEM", "TM", "TP", "TS", "ULTRA"]
#
# # for close filling
# options = Options()
# # options.add_argument('--headless=new')
# options.add_argument("--disable-blink-features=AutomationControlled")
#
# # Exclude the collection of enable-automation switches
# options.add_experimental_option("excludeSwitches", ["enable-automation"])
#
# # Turn-off userAutomationExtension
# options.add_experimental_option("useAutomationExtension", False)
#
# cash_close_list = ['AARTIIND', 'ABB', 'ABCAPITAL', 'ABFRL', 'ADANIENT', 'ADANIPORTS', 'ALKEM', 'AMBUJACEM',
#                      'APOLLOHOSP', 'APOLLOTYRE', 'ASHOKLEY', 'ASTRAL', 'ATUL', 'AUBANK', 'AUROPHARMA', 'BAJAJFINSV',
#                      'BAJFINANCE', 'BALKRISIND', 'BALRAMCHIN', 'BANDHANBNK', 'BANKBARODA', 'BATAINDIA', 'BEL',
#                      'BHARATFORG', 'BIOCON', 'BRITANNIA', 'BSOFT', 'CANBK', 'CANFINHOME', 'CHAMBLFERT', 'CHOLAFIN',
#                      'CIPLA', 'COFORGE', 'CONCOR', 'COROMANDEL', 'CROMPTON', 'CUMMINSIND', 'DABUR', 'DALBHARAT',
#                      'DEEPAKFERT', 'DEEPAKNTR', 'DELTACORP', 'DIVISLAB', 'DIXON', 'DLF', 'DRREDDY', 'ESCORTS',
#                      'EXIDEIND', 'GLENMARK', 'GLS', 'GNFC', 'GODREJCP', 'GODREJPROP', 'GRANULES', 'GRASIM', 'GUJGASLTD',
#                      'HAL', 'HAVELLS', 'HCLTECH', 'HDFCAMC', 'HDFCLIFE', 'HINDALCO', 'HINDCOPPER', 'ICICIGI',
#                      'ICICIPRULI', 'IEX', 'IGL', 'INDHOTEL', 'INDIACEM', 'INDIAMART', 'INDIGO', 'INDUSINDBK',
#                      'INDUSTOWER', 'INTELLECT', 'IPCALAB', 'JINDALSTEL', 'JKCEMENT', 'JSWSTEEL', 'JUBLFOOD',
#                      'KOTAKBANK', 'LALPATHLAB', 'LAURUSLABS', 'LICHSGFIN', 'LTIM', 'LTTS', 'LUPIN', 'M%26MFIN',
#                      'MANAPPURAM', 'MARICO', 'MCDOWELL-N', 'MCX', 'METROPOLIS', 'MFSL', 'MGL', 'MPHASIS', 'MUTHOOTFIN',
#                      'NAM-INDIA', 'NAUKRI', 'NAVINFLUOR', 'NMDC', 'NTPC', 'OBEROIRLTY', 'PEL', 'PERSISTENT', 'PETRONET',
#                      'PIDILITIND', 'POLYCAB', 'POWERGRID', 'RAIN', 'RAMCOCEM', 'RBLBANK', 'RECLTD', 'SBICARD',
#                      'SBILIFE', 'SIEMENS', 'SRF', 'STAR', 'SUNPHARMA', 'SYNGENE', 'TATACOMM', 'TATAMOTORS', 'TECHM',
#                      'TORNTPHARM', 'TORNTPOWER', 'TRENT', 'TVSMOTOR', 'UBL', 'ULTRACEMCO', 'UPL', 'VEDL', 'VOLTAS',
#                      'ZEEL', 'ZYDUSLIFE']
#
# # cash_close_list1 = ["M%26M"]
#
# manual = []         # list to keep track of the shares whose values selenium couldn't get
# close = []
# # close = ['2446.95', '398.5', '1478.0', '7005.0', '227.65', '188.4', '234.15', '470.1', '3337.8', '132.8', '1170.0',
# # '1608.5', '448.85', '959.0', '1387.8', '1392.95', '662.9', '421.45', '1544.95', '1544.95', '2573.2', '561.0', '548.9',
# # '1000.05', '606.8', '230.95', '117.85', '8024.95']
#
# for share in cash_close_list:
#     driver = webdriver.Chrome(options=options)
#
#     driver.get(f"https://www.nseindia.com/get-quotes/equity?symbol={share}")
#
#     try:
#         sleep(2)
#         WebDriverWait(driver, 10).until(ec.presence_of_element_located((By.ID, 'quoteLtp')))
#
#         close_val = driver.find_element(By.ID, "quoteLtp").text
#
#         while close_val == '':
#             driver.refresh()
#             WebDriverWait(driver, 10).until(ec.presence_of_element_located((By.ID, 'quoteLtp')))
#             close_val = driver.find_element(By.ID, "quoteLtp").text
#             sleep(0.5)
#
#         close_val = close_val.replace(",", "")
#
#         # truncating last 0
#         if close_val[len(close_val)-1:len(close_val)] == '0':
#             close_val = close_val[:len(close_val)-1]
#
#         close.append(close_val)
#
#         print(f'{share}: {close_val}')
#         if close_val == '':
#             manual.append(share)
#
#     except TimeoutException:        # added temp fix for timeoutexception, will need to check if it works properly or nah
#         close.append('')
#         manual.append(share)
#         print(f"Loading took too much time for {share}!")
#
#     driver.close()
#
# print(close)
# print(manual)
#
# for i in close:
#     print(i)


# import pyautogui as pg
# import numpy as np
# from time import sleep
# #
# # pg.click((797, 1058))
# # sleep(1)
# # # # Take a screenshot of the specified region
# # region = (929, 212, 6, 8)
# # screenshot = pg.screenshot(region=region)
# # #
# # # # Convert the screenshot to a NumPy array
# # pixels = np.array(screenshot)
# #
# # # Calculate the average color
# # average_color = pixels.mean(axis=(0, 1)).astype(int)
# #
# # # Print the average color as an RGB tuple
# # print(f"Average color: {tuple(average_color)}")
#
#
# pg.click((797, 1058))
# sleep(1)

# def check_change():
#     changed = False  # bool representing if screen has changed from black to showing vwap statistics
#
#     # waiting for screen to change before saving
#     while not changed:
#         # # Take a screenshot of the specified region
#         region = (929, 212, 6, 8)
#         screenshot = pg.screenshot(region=region)
#         pixels = np.array(screenshot)
#         # Calculate the average color
#         average_color = pixels.mean(axis=(0, 1)).astype(int)
#
#         # Print the average color as an RGB tuple
#         print(f"Average color: {tuple(average_color)}")
#
#         if 200 > average_color[1] > 50:  # if color is green (screen has updated)
#             print("Displayed the vwap data")
#             print(f"Average color: {tuple(average_color)}")
#             changed = True
#
#
#
# check_change()


# import os
# from time import sleep
#
# f = ['AARTIIND', 'ABB', 'AMBUJACEM', 'APOLLOHOSP', 'APOLLOTYRE', 'ASHOKLEY', 'AUROPHARMA', 'BANKBARODA', 'BEL',
#      'BSOFT', 'CANBK', 'CANFINHOME', 'CHAMBLFERT', 'CHOLAFIN', 'DIXON', 'DLF', 'ESCORTS', 'EXIDEIND', 'GNFC',
#      'GODREJPROP', 'HAL', 'HAVELLS', 'HCLTECH', 'HINDALCO', 'HINDCOPPER', 'IGL', 'INDIACEM', 'INDUSINDBK', 'INDUSTOWER',
#      'JINDALSTEL', 'JUBLFOOD', 'LALPATHLAB', 'LICHSGFIN', 'LUPIN', 'MANAPPURAM', 'MCX', 'METROPOLIS', 'MGL',
#      'MUTHOOTFIN', 'NAM-INDIA', 'NMDC', 'NTPC', 'PETRONET', 'RAMCOCEM', 'RBLBANK', 'RECLTD', 'STAR', 'TRENT', 'VEDL']
#
#
# for fl in f:
#     path = rf'E:\Daily Data work\ALGORITHM\{fl}.xlsx'
#     os.startfile(path)


import pandas as pd
import re
from date_variables import yr, date, mnth
from zipfile import ZipFile

share_list = ["BANKNIFTY", "NIFTY", "ADANIPORTS", "AUROPHARMA", "CANBK", "DLF", "HINDALCO",
              "ICICIBANK", "JINDALSTEL", "RELIANCE", "SBIN", "TATACONSUM", "TATAMOTORS",
              "TATASTEEL", "TCS", "TITAN"]

md_path_zipped = rf"E:\chrome downloads\BhavCopy_NSE_FO_0_0_0_{yr}{date[3:5]}{date[:2]}_F_0000.csv.zip"     # .zip file path of downloaded cash bhavcopy
md_path = rf"E:\chrome downloads"

# take data from col 22 and col 14 has month as cur month in caps and col 13 is null
#
# # extracting .zip file
# with ZipFile(md_path_zipped, 'r') as zObject:
#     zObject.extractall(path=md_path)
#
md_file_path = rf"E:\Daily Data work\MD files\{yr}\{mnth}\fo{date[:2]}{mnth}20{date[6:]}bhav.xlsx"

df = pd.read_csv(md_path_zipped[:-4])       # removing the .zip extension after unzipping
df = df.drop(df.columns[-9:], axis=1)
df = df.drop(columns=['BizDt', 'Src', 'FinInstrmTp', 'FinInstrmId', 'ISIN', 'SctySrs', 'OpnIntrst', 'ChngInOpnIntrst'])
df2 = df.copy()     # this one will actually be saved as md file
df = df[df['OptnTp'].isnull()]

for share in share_list:
    ltp = df.loc[(df['TckrSymb'] == share) & df['FinInstrmNm'].str.contains(mnth), 'SttlmPric'].iloc[0]
    print(f"{share}: {ltp}")


# BANKNIFTY: 52560.5
# NIFTY: 23638.9
# ADANIPORTS: 1482.5
# AUROPHARMA: 1308.35
# CANBK: 115.45
# DLF: 839.6
# HINDALCO: 698.6
# ICICIBANK: 1234.9
# JINDALSTEL: 1031.3
# RELIANCE: 3208.1
# SBIN: 860.05
# TATACONSUM: 1154.4
# TATAMOTORS: 1006.75
# TATASTEEL: 172.7
# TCS: 3996.8
# TITAN: 3171.35