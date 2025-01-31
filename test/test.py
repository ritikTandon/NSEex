# # # # # from time import sleep
# # # # # from selenium import webdriver
# # # # # from selenium.webdriver.support.ui import WebDriverWait
# # # # # from selenium.webdriver.support import expected_conditions as ec
# # # # # from selenium.webdriver.common.by import By
# # # # # from selenium.webdriver.chrome.options import Options
# # # # # from selenium.common.exceptions import TimeoutException
# # # # #
# # # # # # headers = {'User-Agent': 'Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:52.0) Gecko/20100101 Firefox/52.0'}
# # # # #
# # # # # cash_share_list = ["AARTIIND", "ADANI", "APOLLO", "BAJFINSV", "BAJFIN", "BANBK", "BARODA", "COALIND", "DLF", "EICHER",
# # # # #                    "FEDBANK", "HCL", "HDFC", "HIND", "ICICI", "INDUSIND", "INFY", "JIND", "LIC", "M&M", "M&MFIN", "NTPC",
# # # # #                    "REL", "SBIN", "SUNTV", "TCHEM", "TM", "TP", "TS", "ULTRA"]
# # # # #
# # # # # # for close filling
# # # # # options = Options()
# # # # # # options.add_argument('--headless=new')
# # # # # options.add_argument("--disable-blink-features=AutomationControlled")
# # # # #
# # # # # # Exclude the collection of enable-automation switches
# # # # # options.add_experimental_option("excludeSwitches", ["enable-automation"])
# # # # #
# # # # # # Turn-off userAutomationExtension
# # # # # options.add_experimental_option("useAutomationExtension", False)
# # # # #
# # # # # cash_close_list = ['AARTIIND', 'ABB', 'ABCAPITAL', 'ABFRL', 'ADANIENT', 'ADANIPORTS', 'ALKEM', 'AMBUJACEM',
# # # # #                      'APOLLOHOSP', 'APOLLOTYRE', 'ASHOKLEY', 'ASTRAL', 'ATUL', 'AUBANK', 'AUROPHARMA', 'BAJAJFINSV',
# # # # #                      'BAJFINANCE', 'BALKRISIND', 'BALRAMCHIN', 'BANDHANBNK', 'BANKBARODA', 'BATAINDIA', 'BEL',
# # # # #                      'BHARATFORG', 'BIOCON', 'BRITANNIA', 'BSOFT', 'CANBK', 'CANFINHOME', 'CHAMBLFERT', 'CHOLAFIN',
# # # # #                      'CIPLA', 'COFORGE', 'CONCOR', 'COROMANDEL', 'CROMPTON', 'CUMMINSIND', 'DABUR', 'DALBHARAT',
# # # # #                      'DEEPAKFERT', 'DEEPAKNTR', 'DELTACORP', 'DIVISLAB', 'DIXON', 'DLF', 'DRREDDY', 'ESCORTS',
# # # # #                      'EXIDEIND', 'GLENMARK', 'GLS', 'GNFC', 'GODREJCP', 'GODREJPROP', 'GRANULES', 'GRASIM', 'GUJGASLTD',
# # # # #                      'HAL', 'HAVELLS', 'HCLTECH', 'HDFCAMC', 'HDFCLIFE', 'HINDALCO', 'HINDCOPPER', 'ICICIGI',
# # # # #                      'ICICIPRULI', 'IEX', 'IGL', 'INDHOTEL', 'INDIACEM', 'INDIAMART', 'INDIGO', 'INDUSINDBK',
# # # # #                      'INDUSTOWER', 'INTELLECT', 'IPCALAB', 'JINDALSTEL', 'JKCEMENT', 'JSWSTEEL', 'JUBLFOOD',
# # # # #                      'KOTAKBANK', 'LALPATHLAB', 'LAURUSLABS', 'LICHSGFIN', 'LTIM', 'LTTS', 'LUPIN', 'M%26MFIN',
# # # # #                      'MANAPPURAM', 'MARICO', 'MCDOWELL-N', 'MCX', 'METROPOLIS', 'MFSL', 'MGL', 'MPHASIS', 'MUTHOOTFIN',
# # # # #                      'NAM-INDIA', 'NAUKRI', 'NAVINFLUOR', 'NMDC', 'NTPC', 'OBEROIRLTY', 'PEL', 'PERSISTENT', 'PETRONET',
# # # # #                      'PIDILITIND', 'POLYCAB', 'POWERGRID', 'RAIN', 'RAMCOCEM', 'RBLBANK', 'RECLTD', 'SBICARD',
# # # # #                      'SBILIFE', 'SIEMENS', 'SRF', 'STAR', 'SUNPHARMA', 'SYNGENE', 'TATACOMM', 'TATAMOTORS', 'TECHM',
# # # # #                      'TORNTPHARM', 'TORNTPOWER', 'TRENT', 'TVSMOTOR', 'UBL', 'ULTRACEMCO', 'UPL', 'VEDL', 'VOLTAS',
# # # # #                      'ZEEL', 'ZYDUSLIFE']
# # # # #
# # # # # # cash_close_list1 = ["M%26M"]
# # # # #
# # # # # manual = []         # list to keep track of the shares whose values selenium couldn't get
# # # # # close = []
# # # # # # close = ['2446.95', '398.5', '1478.0', '7005.0', '227.65', '188.4', '234.15', '470.1', '3337.8', '132.8', '1170.0',
# # # # # # '1608.5', '448.85', '959.0', '1387.8', '1392.95', '662.9', '421.45', '1544.95', '1544.95', '2573.2', '561.0', '548.9',
# # # # # # '1000.05', '606.8', '230.95', '117.85', '8024.95']
# # # # #
# # # # # for share in cash_close_list:
# # # # #     driver = webdriver.Chrome(options=options)
# # # # #
# # # # #     driver.get(f"https://www.nseindia.com/get-quotes/equity?symbol={share}")
# # # # #
# # # # #     try:
# # # # #         sleep(2)
# # # # #         WebDriverWait(driver, 10).until(ec.presence_of_element_located((By.ID, 'quoteLtp')))
# # # # #
# # # # #         close_val = driver.find_element(By.ID, "quoteLtp").text
# # # # #
# # # # #         while close_val == '':
# # # # #             driver.refresh()
# # # # #             WebDriverWait(driver, 10).until(ec.presence_of_element_located((By.ID, 'quoteLtp')))
# # # # #             close_val = driver.find_element(By.ID, "quoteLtp").text
# # # # #             sleep(0.5)
# # # # #
# # # # #         close_val = close_val.replace(",", "")
# # # # #
# # # # #         # truncating last 0
# # # # #         if close_val[len(close_val)-1:len(close_val)] == '0':
# # # # #             close_val = close_val[:len(close_val)-1]
# # # # #
# # # # #         close.append(close_val)
# # # # #
# # # # #         print(f'{share}: {close_val}')
# # # # #         if close_val == '':
# # # # #             manual.append(share)
# # # # #
# # # # #     except TimeoutException:        # added temp fix for timeoutexception, will need to check if it works properly or nah
# # # # #         close.append('')
# # # # #         manual.append(share)
# # # # #         print(f"Loading took too much time for {share}!")
# # # # #
# # # # #     driver.close()
# # # # #
# # # # # print(close)
# # # # # print(manual)
# # # # #
# # # # # for i in close:
# # # # #     print(i)
# # # #
# # # #
# # # # # import pyautogui as pg
# # # # # import numpy as np
# # # # # from time import sleep
# # # # # #
# # # # # # pg.click((797, 1058))
# # # # # # sleep(1)
# # # # # # # # Take a screenshot of the specified region
# # # # # # region = (929, 212, 6, 8)
# # # # # # screenshot = pg.screenshot(region=region)
# # # # # # #
# # # # # # # # Convert the screenshot to a NumPy array
# # # # # # pixels = np.array(screenshot)
# # # # # #
# # # # # # # Calculate the average color
# # # # # # average_color = pixels.mean(axis=(0, 1)).astype(int)
# # # # # #
# # # # # # # Print the average color as an RGB tuple
# # # # # # print(f"Average color: {tuple(average_color)}")
# # # # #
# # # # #
# # # # # pg.click((797, 1058))
# # # # # sleep(1)
# # # #
# # # # # def check_change():
# # # # #     changed = False  # bool representing if screen has changed from black to showing vwap statistics
# # # # #
# # # # #     # waiting for screen to change before saving
# # # # #     while not changed:
# # # # #         # # Take a screenshot of the specified region
# # # # #         region = (929, 212, 6, 8)
# # # # #         screenshot = pg.screenshot(region=region)
# # # # #         pixels = np.array(screenshot)
# # # # #         # Calculate the average color
# # # # #         average_color = pixels.mean(axis=(0, 1)).astype(int)
# # # # #
# # # # #         # Print the average color as an RGB tuple
# # # # #         print(f"Average color: {tuple(average_color)}")
# # # # #
# # # # #         if 200 > average_color[1] > 50:  # if color is green (screen has updated)
# # # # #             print("Displayed the vwap data")
# # # # #             print(f"Average color: {tuple(average_color)}")
# # # # #             changed = True
# # # # #
# # # # #
# # # # #
# # # # # check_change()
# # # #
# # # #
# # # # # import os
# # # # # from time import sleep
# # # # #
# # # # # f = ['AARTIIND', 'ABB', 'AMBUJACEM', 'APOLLOHOSP', 'APOLLOTYRE', 'ASHOKLEY', 'AUROPHARMA', 'BANKBARODA', 'BEL',
# # # # #      'BSOFT', 'CANBK', 'CANFINHOME', 'CHAMBLFERT', 'CHOLAFIN', 'DIXON', 'DLF', 'ESCORTS', 'EXIDEIND', 'GNFC',
# # # # #      'GODREJPROP', 'HAL', 'HAVELLS', 'HCLTECH', 'HINDALCO', 'HINDCOPPER', 'IGL', 'INDIACEM', 'INDUSINDBK', 'INDUSTOWER',
# # # # #      'JINDALSTEL', 'JUBLFOOD', 'LALPATHLAB', 'LICHSGFIN', 'LUPIN', 'MANAPPURAM', 'MCX', 'METROPOLIS', 'MGL',
# # # # #      'MUTHOOTFIN', 'NAM-INDIA', 'NMDC', 'NTPC', 'PETRONET', 'RAMCOCEM', 'RBLBANK', 'RECLTD', 'STAR', 'TRENT', 'VEDL']
# # # # #
# # # # #
# # # # # for fl in f:
# # # # #     path = rf'E:\Daily Data work\ALGORITHM\{fl}.xlsx'
# # # # #     os.startfile(path)
# # # #
# # # #
# # # # import pandas as pd
# # # # import re
# # # # from date_variables import yr, date, mnth
# # # # from zipfile import ZipFile
# # # #
# # # # share_list = ["BANKNIFTY", "NIFTY"]
# # # #
# # # # md_path_zipped = rf"E:\chrome downloads\BhavCopy_NSE_FO_0_0_0_{yr}{date[3:5]}{date[:2]}_F_0000.csv.zip"     # .zip file path of downloaded cash bhavcopy
# # # # md_path = rf"E:\chrome downloads"
# # # #
# # # # # take data from col 22 and col 14 has month as cur month in caps and col 13 is null
# # # # #
# # # # # # extracting .zip file
# # # # # with ZipFile(md_path_zipped, 'r') as zObject:
# # # # #     zObject.extractall(path=md_path)
# # # # #
# # # # md_file_path = rf"E:\Daily Data work\MD files\{yr}\{mnth}\fo{date[:2]}{mnth}20{date[6:]}bhav.xlsx"
# # # #
# # # # df = pd.read_csv(md_path_zipped[:-4])       # removing the .zip extension after unzipping
# # # # df = df.drop(df.columns[-9:], axis=1)
# # # # df = df.drop(columns=['BizDt', 'Src', 'FinInstrmTp', 'FinInstrmId', 'ISIN', 'SctySrs', 'OpnIntrst', 'ChngInOpnIntrst'])
# # # # df2 = df.copy()     # this one will actually be saved as md file
# # # # df = df[df['OptnTp'].isnull()]
# # # #
# # # # import datetime
# # # # import calendar
# # # # from date_variables import date,mnth, yr
# # # #
# # # #
# # # # def checkExpiry(day, month_abbr, year):
# # # #     # Mapping of month abbreviations to month numbers
# # # #     month_map = {
# # # #         "JAN": 1, "FEB": 2, "MAR": 3, "APR": 4, "MAY": 5, "JUN": 6,
# # # #         "JUL": 7, "AUG": 8, "SEP": 9, "OCT": 10, "NOV": 11, "DEC": 12
# # # #     }
# # # #
# # # #     # Convert the month abbreviation to a month number
# # # #     month = month_map.get(month_abbr.upper())
# # # #
# # # #     if month is None:
# # # #         raise ValueError("Invalid month abbreviation")
# # # #
# # # #     # Get the number of days in the month
# # # #     _, last_day_of_month = calendar.monthrange(year, month)
# # # #
# # # #     # Find the last Thursday of the month
# # # #     last_date_of_month = datetime.date(year, month, last_day_of_month)
# # # #     while last_date_of_month.weekday() != 3:  # 3 corresponds to Thursday
# # # #         last_date_of_month -= datetime.timedelta(days=1)
# # # #
# # # #     # Check if the given day is greater than or equal to the last Thursday
# # # #     given_date = datetime.date(year, month, day)
# # # #     if given_date >= last_date_of_month:
# # # #         # Move to the next month
# # # #         if month == 12:
# # # #             next_month = "JAN"
# # # #         else:
# # # #             next_month = list(month_map.keys())[list(month_map.values()).index(month) + 1]
# # # #         return next_month
# # # #     else:
# # # #         return month_abbr
# # # #
# # # #
# # # # d = int(date[:2])
# # # # m = mnth
# # # # y = int(yr)
# # # #
# # # #
# # # # for share in share_list:
# # # #     ltp = df.loc[(df['TckrSymb'] == share) & df['FinInstrmNm'].str.contains(checkExpiry(d, m, y)), 'SttlmPric'].iloc[0]
# # # #     print(f"{share}: {ltp}")
# # # #
# # # #
# # # #
# # # # # BANKNIFTY: 52560.5
# # # # # NIFTY: 23638.9
# # # # # ADANIPORTS: 1482.5
# # # # # AUROPHARMA: 1308.35
# # # # # CANBK: 115.45
# # # # # DLF: 839.6
# # # # # HINDALCO: 698.6
# # # # # ICICIBANK: 1234.9
# # # # # JINDALSTEL: 1031.3
# # # # # RELIANCE: 3208.1
# # # # # SBIN: 860.05
# # # # # TATACONSUM: 1154.4
# # # # # TATAMOTORS: 1006.75
# # # # # TATASTEEL: 172.7
# # # # # TCS: 3996.8
# # # # # TITAN: 3171.35
# # # import re
# # # # //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
# # # # 1160, 233, 57, 56
# # # # 1157-1161, 1157-1176, 1156-1175
# # #
# # # # hl, hl%, cl, cl%, trend
# # #
# # # import time
# # # import openpyxl as xl
# # # import pyautogui as pg
# # # from time import sleep
# # #
# # # from openpyxl.formatting.rule import CellIsRule
# # # from selenium.webdriver.support.ui import WebDriverWait
# # # from selenium.webdriver.support import expected_conditions as ec
# # # from selenium.webdriver.common.by import By
# # # from selenium.webdriver.chrome.options import Options
# # # from selenium.common.exceptions import TimeoutException
# # # from selenium import webdriver
# # # from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
# # #
# # # algo_share_list = ['AARTIIND', '02 ABB', 'ABCAPITAL', 'ABFRL', 'ADANIENT', 'ADANIPORTS', 'ALKEM', 'AMBUJACEM',
# # #                          'APOLLOHOSP', 'APOLLOTYRE', '03 ASHOKLEY', 'ASTRAL', 'ATUL', 'AUBANK', 'AUROPHARMA', 'BAJAJFINSV',
# # #                          'BAJFINANCE', 'BALKRISIND', 'BALRAMCHIN', 'BANDHANBNK', 'BANKBARODA', 'BATAINDIA', 'BEL',
# # #                          'BHARATFORG', '04 BHEL', 'BIOCON', 'BRITANNIA', 'BSOFT', 'CANBK', 'CANFINHOME', 'CHAMBLFERT', 'CHOLAFIN',
# # #                          'CIPLA', 'COFORGE', 'CONCOR', 'COROMANDEL', 'CROMPTON', 'CUMMINSIND', 'DABUR', 'DALBHARAT',
# # #                          'DEEPAKFERT', 'DEEPAKNTR', 'DELTACORP', 'DIVISLAB', '05 DIXON', 'DLF', 'DRREDDY', 'ESCORTS',
# # #                          'EXIDEIND', 'GLENMARK', 'GLS', 'GNFC', 'GODREJCP', 'GODREJPROP', 'GRANULES', 'GRASIM', 'GUJGASLTD',
# # #                          'HAL', 'HAVELLS', 'HCLTECH', 'HDFCAMC', 'HDFCLIFE', 'HINDALCO', 'HINDCOPPER', 'ICICIGI',
# # #                          'ICICIPRULI', 'IEX', 'IGL', 'INDHOTEL', 'INDIACEM', 'INDIAMART', 'INDIGO', 'INDUSINDBK',
# # #                          'INDUSTOWER', 'INTELLECT', 'IPCALAB', 'JINDALSTEL', 'JKCEMENT', 'JSWSTEEL', 'JUBLFOOD',
# # #                          'KOTAKBANK', 'LALPATHLAB', 'LAURUSLABS', 'LICHSGFIN', 'LTIM', 'LTTS', 'LUPIN', 'M%26MFIN',
# # #                          'MANAPPURAM', 'MARICO', 'MCDOWELL-N', 'MCX', 'METROPOLIS', 'MFSL', 'MGL', 'MPHASIS', 'MUTHOOTFIN',
# # #                          'NAM-INDIA', 'NAUKRI', 'NAVINFLUOR', 'NMDC', 'NTPC', 'OBEROIRLTY', '09 ONGC', 'PEL', 'PERSISTENT', 'PETRONET',
# # #                          'PIDILITIND', 'POLYCAB', 'POWERGRID', 'RAIN', 'RAMCOCEM', 'RBLBANK', '10 RECLTD', 'SBICARD',
# # #                          'SBILIFE', 'SIEMENS', 'SRF', 'STAR', 'SUNPHARMA', 'SYNGENE', 'TATACOMM', 'TATAMOTORS', 'TECHM',
# # #                          'TORNTPHARM', 'TORNTPOWER', 'TRENT', 'TVSMOTOR', 'UBL', 'ULTRACEMCO', 'UPL', 'VEDL', 'VOLTAS',
# # #                          'ZEEL', 'ZYDUSLIFE']
# # #
# # # daily_cell_range = "H3:K3000"
# # # weekly_cell_range = "E4:I1000"
# # # other_cell_range = "F4:J1000"
# # #
# # # font = Font(name='Arial', size=11, bold=True)  # text font
# # # heading_font = Font(name='Arial', size=11, bold=True, color="FFFFFF")  # heading
# # # fill = PatternFill(patternType='solid', fgColor="0000ff")  # blue
# # # red = Font("Arial", 11, color='ff0000', bold=True)
# # # blue = Font("Arial", 11, color="0000ff", bold=True)
# # # alignment = Alignment(horizontal='center')
# # # border = Border(
# # #     left=Side(style='thin'),
# # #     right=Side(style='thin'),
# # #     top=Side(style='thin'),
# # #     bottom=Side(style='thin')
# # # )
# # #
# # # red_text_rule = CellIsRule(operator='lessThan', formula=['0'], font=red)
# # # blue_text_rule = CellIsRule(operator='greaterThanOrEqual', formula=['0'], font=blue)
# # #
# # # sh_row = 2
# # # d_row = 1101
# # # w_row = 221
# # # m_row = 54
# # # cl_row = 53
# # # w_range = [1097, 1101]
# # # m_range = [1091, 1110]
# # # cl_range = [1091, 1110]
# # # sh_count = 0
# # #
# # # xp_high = '/html/body/div[11]/div/div/section/div/div/div/div/div/div[2]/div/div/div/div[2]/div/div[1]/section/div/div[3]/div/table/tbody/tr[2]/td[4]'
# # # xp_low = '/html/body/div[11]/div/div/section/div/div/div/div/div/div[2]/div/div/div/div[2]/div/div[1]/section/div/div[3]/div/table/tbody/tr[2]/td[5]'
# # #
# # # options = Options()
# # # # options.add_argument('--headless=new')
# # # options.add_argument("--disable-blink-features=AutomationControlled")
# # # options.add_argument("--start-maximized")
# # #
# # # # Exclude the collection of enable-automation switches
# # # options.add_experimental_option("excludeSwitches", ["enable-automation"])
# # #
# # # # Turn-off userAutomationExtension
# # # options.add_experimental_option("useAutomationExtension", False)
# # #
# # # algo_close_list = ['AARTIIND', 'ABB', 'ABCAPITAL', 'ABFRL', 'ADANIENT', 'ADANIPORTS', 'ALKEM', 'AMBUJACEM',
# # #                      'APOLLOHOSP', 'APOLLOTYRE', 'ASHOKLEY', 'ASTRAL', 'ATUL', 'AUBANK', 'AUROPHARMA', 'BAJAJ-AUTO', 'BAJAJFINSV',
# # #                      'BAJFINANCE', 'BALKRISIND', 'BALRAMCHIN', 'BANDHANBNK', 'BANKBARODA', 'BATAINDIA', 'BEL',
# # #                      'BHARATFORG', 'BHEL', 'BIOCON', 'BRITANNIA', 'BSOFT', 'CANBK', 'CANFINHOME', 'CHAMBLFERT', 'CHOLAFIN',
# # #                      'CIPLA', 'COFORGE', 'CONCOR', 'COROMANDEL', 'CROMPTON', 'CUMMINSIND', 'DABUR', 'DALBHARAT',
# # #                      'DEEPAKFERT', 'DEEPAKNTR', 'DELTACORP', 'DIVISLAB', 'DIXON', 'DLF', 'DRREDDY', 'ESCORTS',
# # #                      'EXIDEIND', 'GLENMARK', 'GLS', 'GNFC', 'GODREJCP', 'GODREJPROP', 'GRANULES', 'GRASIM', 'GUJGASLTD',
# # #                      'HAL', 'HAVELLS', 'HCLTECH', 'HDFCAMC', 'HDFCLIFE', 'HINDALCO', 'HINDCOPPER', 'ICICIGI',
# # #                      'ICICIPRULI', 'IEX', 'IGL', 'INDHOTEL', 'INDIACEM', 'INDIAMART', 'INDIGO', 'INDUSINDBK',
# # #                      'INDUSTOWER', 'INTELLECT', 'IPCALAB', 'JINDALSTEL', 'JKCEMENT', 'JSWSTEEL', 'JUBLFOOD',
# # #                      'KOTAKBANK', 'LALPATHLAB', 'LAURUSLABS', 'LICHSGFIN', 'LTIM', 'LTTS', 'LUPIN', 'M%26MFIN',
# # #                      'MANAPPURAM', 'MARICO', 'UNITDSPR', 'MCX', 'METROPOLIS', 'MFSL', 'MGL', 'MPHASIS', 'MUTHOOTFIN',
# # #                      'NAM-INDIA', 'NAUKRI', 'NAVINFLUOR', 'NMDC', 'NTPC', 'OBEROIRLTY', 'ONGC', 'PEL', 'PERSISTENT', 'PETRONET',
# # #                      'PIDILITIND', 'POLYCAB', 'POWERGRID', 'RAIN', 'RAMCOCEM', 'RBLBANK', 'RECLTD', 'SBICARD',
# # #                      'SBILIFE', 'SIEMENS', 'SRF', 'STAR', 'SUNPHARMA', 'SYNGENE', 'TATACOMM', 'TATAMOTORS', 'TCS', 'TECHM',
# # #                      'TITAN', 'TORNTPHARM', 'TORNTPOWER', 'TRENT', 'TVSMOTOR', 'UBL', 'ULTRACEMCO', 'UPL', 'VEDL', 'VOLTAS',
# # #                      'ZEEL', 'ZYDUSLIFE']
# # #
# # #
# # # doub = ['ADANIPORTS', 'CHOLAFIN']
# # # add = 0
# # # sh_row = 34
# # # for share in algo_close_list:
# # #     # if share takes 2 line as name
# # #     if share in doub:
# # #         add = 20
# # #     left = (1105, 530+add)
# # #     ll = (888, 575+add)  # 6
# # #     ldate = (1052, 701+add)
# # #     hist_data = (593, 439+add)
# # #     right = (1387, 532+add)
# # #     rl = (1172, 573+add)  # 7
# # #     rdate = (1334, 702+add)
# # #     flter = (1467, 527+add)
# # #     hlwb = xl.load_workbook(rf'C:\Users\admin\PycharmProjects\daily data\test\hl_test.xlsx')
# # #     hlsh = hlwb['Sheet2']
# # #     driver = webdriver.Chrome(options=options)
# # #
# # #     driver.get(f"https://www.nseindia.com/get-quotes/equity?symbol={share}")
# # #
# # #
# # #
# # #     try:
# # #         myElem = WebDriverWait(driver, 20).until(ec.presence_of_element_located((By.ID, 'historic_data')))
# # #         sleep(2)
# # #         pg.click(hist_data)
# # #         sleep(2)
# # #         # pg.click(left)
# # #         # sleep(2)
# # #         # pg.click(ll, clicks=6, interval=0.6)
# # #         # sleep(2)
# # #         # pg.click(ldate)
# # #         # sleep(2)
# # #         # pg.click(right)
# # #         # sleep(2)
# # #         # pg.click(rl, clicks=7, interval=0.6)
# # #         # sleep(2)
# # #         # pg.click(rdate)
# # #         # sleep(2)
# # #         # pg.click(flter)
# # #         # sleep(2)
# # #         he = WebDriverWait(driver, 20).until(ec.presence_of_element_located((By.XPATH, xp_high)))
# # #         high = float(driver.find_element(By.XPATH, xp_high).text.replace(",", ""))
# # #         low = float(driver.find_element(By.XPATH, xp_low).text.replace(",", ""))
# # #
# # #         hlsh.cell(sh_row, 2).value = high
# # #         hlsh.cell(sh_row, 3).value = low
# # #         sh_row += 1
# # #
# # #         hlwb.save(rf'C:\Users\admin\PycharmProjects\daily data\test\hl_test.xlsx')
# # #         print(f"{high} {low} {share}")
# # #
# # #     except TimeoutException:
# # #         print(f"Loading took too much time for {share}!")
# # #
# # #     driver.close()
# # #
# # #
# # # # # copying high and low from algo hl sheet (after putting correct high and low values for that date)
# # # # for share in algo_share_list:
# # # #     hlwb = xl.load_workbook(rf'C:\Users\admin\PycharmProjects\daily data\test\hl_test.xlsx')
# # # #     hlsh = hlwb['Sheet2']
# # # #
# # # #     wb = xl.load_workbook(rf'E:\Daily Data work\ALGORITHM\ALGORITHM OLD\{share}.xlsx')
# # # #     sh = wb['D']
# # # #
# # # #     sh.cell(d_row, 2).value = hlsh.cell(sh_row, 2).value        # high copy
# # # #     sh.cell(d_row, 3).value = hlsh.cell(sh_row, 3).value        # low copy
# # # #
# # # #     wb.save(rf'C:\Users\admin\PycharmProjects\daily data\test\algo\{share}.xlsx')
# # # #     sh_row += 1
# # # #     sh_count += 1
# # # #
# # # #     print(f"{share} done")
# # # #
# # # # print(f"{sh_count} shares done")
# # #
# # #
# # # # for share in algo_share_list:
# # # #     wb = xl.load_workbook(rf'E:\Daily Data work\ALGORITHM\ALGORITHM OLD\{share}.xlsx')
# # # #     d = wb['D']
# # # #     w = wb['W']
# # # #     m = wb['M']
# # # #     cl = wb['Cl']
# # # #
# # # #     high = 0
# # # #     low = 9999999
# # # #
# # # #     # weekly
# # # #     for row in range(w_range[0], w_range[1]+1):
# # # #         high_cell = d.cell(row, 2)
# # # #         low_cell = d.cell(row, 3)
# # # #
# # # #         if high_cell.value is not None and high_cell.value > high:
# # # #             high = high_cell.value
# # # #
# # # #         if low_cell.value is not None and low_cell.value < low and low_cell.value != 0:
# # # #             low = low_cell.value
# # # #
# # # #     w.cell(w_row, 2).value = high
# # # #     w.cell(w_row, 3).value = low
# # # #
# # # #     high = 0
# # # #     low = 9999999
# # # #
# # # #     # weekly
# # # #     for row in range(m_range[0], m_range[1] + 1):
# # # #         high_cell = d.cell(row, 2)
# # # #         low_cell = d.cell(row, 3)
# # # #
# # # #         if high_cell.value is not None and high_cell.value > high:
# # # #             high = high_cell.value
# # # #
# # # #         if low_cell.value is not None and low_cell.value < low and low_cell.value != 0:
# # # #             low = low_cell.value
# # # #
# # # #     m.cell(m_row, 3).value = high
# # # #     m.cell(m_row, 4).value = low
# # # #
# # # #     high = 0
# # # #     low = 9999999
# # # #
# # # #     # closing
# # # #     for row in range(cl_range[0], cl_range[1] + 1):
# # # #         high_cell = d.cell(row, 2)
# # # #         low_cell = d.cell(row, 3)
# # # #
# # # #         if high_cell.value is not None and high_cell.value > high:
# # # #             high = high_cell.value
# # # #
# # # #         if low_cell.value is not None and low_cell.value < low and low_cell.value != 0:
# # # #             low = low_cell.value
# # # #
# # # #     cl.cell(cl_row, 3).value = high
# # # #     cl.cell(cl_row, 4).value = low
# # # #
# # # #     wb.save(rf'C:\Users\admin\PycharmProjects\daily data\test\algo\{share}.xlsx')
# # # #     print(f'{share} done')
# # #
# # # # for share in algo_share_list:
# # # #     wb = xl.load_workbook(rf'E:\Daily Data work\ALGORITHM\ALGORITHM OLD\{share}.xlsx')
# # # #     sh = wb['D']
# # # #
# # # #     h = sh.cell(mar_15, 2).value
# # # #     l = sh.cell(mar_15, 3).value
# # # #     c = sh.cell(mar_15, 4).value
# # # #     ltp = sh.cell(mar_15, 5).value
# # # #
# # # #     if h > c > l and h > ltp > l:
# # # #         print(f"{share} good!")
# # # #     else:
# # # #         print(f"{share} bad!")
# # #
# # #
# # # # for share in algo_share_list:
# # # #     wb1 = xl.load_workbook(rf'E:\Daily Data work\ALGORITHM\{share}.xlsx')
# # # #     d = wb1['D']
# # # #     w = wb1['W']
# # # #     m = wb1['M']
# # # #     cl = wb1['Cl']
# # # #
# # # #     # headings daily
# # # #     d.cell(2, 8).value = "H/L DIFF"
# # # #     d.cell(2, 9).value = "H/L %"
# # # #     d.cell(2, 10).value = "LTP DIFF"
# # # #     d.cell(2, 11).value = "LTP %"
# # # #
# # # #     # headings weekly
# # # #     w.cell(3, 5).value = "H/L DIFF"
# # # #     w.cell(3, 6).value = "H/L %"
# # # #     w.cell(3, 7).value = "LTP DIFF"
# # # #     w.cell(3, 8).value = "LTP %"
# # # #     w.cell(3, 9).value = "TREND"
# # # #
# # # #     # headings monthly
# # # #     m.cell(3, 6).value = "H/L DIFF"
# # # #     m.cell(3, 7).value = "H/L %"
# # # #     m.cell(3, 8).value = "LTP DIFF"
# # # #     m.cell(3, 9).value = "LTP %"
# # # #     m.cell(3, 10).value = "TREND"
# # # #
# # # #     # headings closing
# # # #     cl.cell(3, 6).value = "H/L DIFF"
# # # #     cl.cell(3, 7).value = "H/L %"
# # # #     cl.cell(3, 8).value = "LTP DIFF"
# # # #     cl.cell(3, 9).value = "LTP %"
# # # #     cl.cell(3, 10).value = "TREND"
# # # #
# # # #     # daily pattern filling
# # # #     for row in range(2, 3000):
# # # #         for col in range(8, 12):
# # # #             cell = d.cell(row, col)
# # # #             cell.font = font
# # # #             cell.alignment = alignment
# # # #             cell.number_format = '0.##'
# # # #
# # # #             if row >= 4:
# # # #                 d.cell(row, 8).value = f'=IF(B{row}="", "", B{row}-C{row})'    # hl diff
# # # #                 d.cell(row, 9).value = f'=IF(B{row}="", "",H{row}/E{row}*100)'    # %
# # # #                 d.cell(row, 10).value = f'=IF(B{row}="", "",IF(E{row-1}="", IF(E{row-2}="", E{row}-E{row-3}, E{row}-E{row-2}), E{row}-E{row-1}))'    # ltp diff
# # # #                 d.cell(row, 11).value = f'=IF(B{row}="", "", J{row}*100/(IF(E{row-1}="", IF(E{row-2}="", E{row-3}, E{row-2}), E{row-1})))'    # %
# # # #
# # # #             if row < 3:
# # # #                 cell.fill = fill
# # # #                 cell.font = heading_font
# # # #                 if col > 9:
# # # #                     d.cell(1, col).fill = fill
# # # #
# # # #     d.conditional_formatting.add(daily_cell_range, red_text_rule)
# # # #     d.conditional_formatting.add(daily_cell_range, blue_text_rule)
# # # #
# # # #     # # weekly pattern filling
# # # #     # for row in range(3, 1001):
# # # #     #     for col in range(5, 10):
# # # #     #         cell = w.cell(row, col)
# # # #     #         cell.font = font
# # # #     #         cell.alignment = alignment
# # # #     #         cell.number_format = '0.##'
# # # #     #         cell.border = border
# # # #     #
# # # #     #         if row >= 4:
# # # #     #             w.cell(row, 5).value = f'=B{row}-C{row}'  # hl diff
# # # #     #             w.cell(row, 6).value = f'=E{row}/D{row}*100'  # %
# # # #     #             w.cell(row, 7).value = f'=D{row}-D{row-1}'  # ltp diff
# # # #     #             w.cell(row, 8).value = f'=G{row}*100/D{row-1}'  # %
# # # #     #
# # # #     # w.conditional_formatting.add(weekly_cell_range, red_text_rule)
# # # #     # w.conditional_formatting.add(weekly_cell_range, blue_text_rule)
# # # #     #
# # # #     # # monthly and cl pattern filling
# # # #     # for row in range(3, 1001):
# # # #     #     for col in range(6, 11):
# # # #     #         cell = m.cell(row, col)
# # # #     #         cell.font = font
# # # #     #         cell.alignment = alignment
# # # #     #         cell.number_format = '0.##'
# # # #     #         cell.border = border
# # # #     #
# # # #     #         cell2 = cl.cell(row, col)
# # # #     #         cell2.font = font
# # # #     #         cell2.alignment = alignment
# # # #     #         cell2.number_format = '0.##'
# # # #     #         cell2.border = border
# # # #     #
# # # #     #         if row >= 4:
# # # #     #             m.cell(row, 5).value = f'=C{row}-D{row}'  # hl diff
# # # #     #             m.cell(row, 6).value = f'=E{row}/D{row}*100'  # %
# # # #     #             m.cell(row, 7).value = f'=D{row}-D{row - 1}'  # ltp diff
# # # #     #             m.cell(row, 8).value = f'=G{row}*100/D{row - 1}'  # %
# # # #     #
# # # #     #             cl.cell(row, 5).value = f'=C{row}-D{row}'  # hl diff
# # # #     #             cl.cell(row, 6).value = f'=E{row}/D{row}*100'  # %
# # # #     #             cl.cell(row, 7).value = f'=D{row}-D{row - 1}'  # ltp diff
# # # #     #             cl.cell(row, 8).value = f'=G{row}*100/D{row - 1}'  # %
# # # #     #
# # # #     # m.conditional_formatting.add(other_cell_range, red_text_rule)
# # # #     # m.conditional_formatting.add(other_cell_range, blue_text_rule)
# # # #     # cl.conditional_formatting.add(other_cell_range, red_text_rule)
# # # #     # cl.conditional_formatting.add(other_cell_range, blue_text_rule)
# # # #
# # # #     print(f'{share} done')
# # # #     sh_count += 1
# # # #     wb1.save(rf'C:\Users\admin\PycharmProjects\daily data\test\algo\{share}.xlsx')
# # # #
# # # # print(sh_count)
# # # #
# # # #
# # # # # def apply_borders(ws, cell_range):
# # # # #     for row in ws[cell_range]:
# # # # #         for cell in row:
# # # # #             cell.border = border
# # # # #
# # # # #
# # # # # for share in algo_share_list:
# # # # #     wb1 = xl.load_workbook(rf'E:\Daily Data work\ALGORITHM\{share}.xlsx')
# # # # #     d = wb1['D']
# # # # #     w = wb1['W']
# # # # #     m = wb1['M']
# # # # #     cl = wb1['Cl']
# # # # #
# # # # #     apply_borders(w, weekly_cell_range)
# # # # #     apply_borders(m, other_cell_range)
# # # # #     apply_borders(cl, other_cell_range)
# # # # #
# # # # #     print(f'{share} done')
# # # # #     wb1.save(rf'C:\Users\admin\PycharmProjects\daily data\test\algo\{share}.xlsx')
# # #
# # #
# # #
# # #
# # #
# # #
# # #
# # #
# # import datetime
# #
# #
# # def get_last_row(sheet, empty=True):
# #     row = sheet.max_row
# #
# #     while True:
# #         if sheet.cell(row, 1).value is not None:
# #             if empty:
# #                 return row+1
# #             else:
# #                 return row
# #
# #         row -= 1
# #
# #
# # import openpyxl as xl
# #
# #
# # cash_shares = {'AARTIIND': 947, 'ADANIENT': 1579, 'APOLLOTYRE': 2946, 'BAJAJFINSERV': 1579, 'BAJAJFINANCE': 1579,
# #                'BANDHANBANK': 1579, 'BANKBARODA': 1579, 'COAL INDIA': 3232, '06 DLF CHL': 4058, 'EICHERMOTOR': 2715,
# #                'FEDRAL BANK': 1579, 'HCLTECH': 1579, 'HDFC': 3936, 'HINDALCO': 947, 'ICICIBANK': 1579,
# #                'INDUSINDBANK': 1579,
# #                'INFY': 2765, 'JINDALS chl': 5195, 'LICHSGFIN': 1579, 'M&M': 1579, '07 M&MFINANCE': 1579,
# #                '08 NTPC': 947, 'RELIANCE CHL': 4793, 'SBIN CHL': 4860, 'SUNTV': 1579, 'TATACHEM': 1579,
# #                '11 TATAMOTOR CHL': 4434, '12 TATAPOWER': 1579, '13 TATASTEEL chl': 4570, 'ULTRACHEM': 2696}
# #
# #
# # # for share in cash_shares:
# # #     path = rf'E:\Daily Data work\CASH\{share}.xlsx'
# # #
# # #     wb = xl.load_workbook(path)
# # #     sheet = wb['Cl']
# # #
# # #     print(f"{share} {sheet.cell(get_last_row(sheet, empty=False), 1).value}")
# # #     # print(sheet.cell(len(sheet['B']), 1).value)
# #
# #
# # dates = """28-12-24 TO 30-01-25
# # 01-02-25 TO 27-02-25
# # 01-03-25 TO 27-03-25
# # 29-03-25 TO 24-04-25
# # 26-04-25 TO 29-05-25
# # 31-05-25 TO 26-06-25
# # 28-06-25 TO 31-07-25
# # 02-08-25 TO 28-08-25
# # 30-08-25 TO 25-09-25
# # 27-09-25 TO 30-10-25
# # 01-11-25 TO 27-11-25
# # 29-11-25 TO 25-12-25
# # 27-12-25 TO 29-01-26
# # 31-01-26 TO 26-02-26
# # 28-02-26 TO 26-03-26
# # 28-03-26 TO 30-04-26
# # 02-05-26 TO 28-05-26
# # 30-05-26 TO 25-06-26
# # 27-06-26 TO 30-07-26
# # 01-08-26 TO 27-08-26
# # 29-08-26 TO 24-09-26
# # 26-09-26 TO 29-10-26
# # 31-10-26 TO 26-11-26
# # 28-11-26 TO 31-12-26
# # 02-01-27 TO 28-01-27
# # 30-01-27 TO 25-02-27
# # 27-02-27 TO 25-03-27
# # 27-03-27 TO 29-04-27
# # 01-05-27 TO 27-05-27
# # 29-05-27 TO 24-06-27
# # 26-06-27 TO 29-07-27
# # 31-07-27 TO 26-08-27
# # 28-08-27 TO 30-09-27
# # 02-10-27 TO 28-10-27
# # 30-10-27 TO 25-11-27
# # 27-11-27 TO 30-12-27
# # 01-01-28 TO 27-01-28
# # 29-01-28 TO 24-02-28
# # 26-02-28 TO 30-03-28
# # 01-04-28 TO 27-04-28
# # 29-04-28 TO 25-05-28
# # 27-05-28 TO 29-06-28
# # 01-07-28 TO 27-07-28
# # 29-07-28 TO 31-08-28
# # 02-09-28 TO 28-09-28
# # 30-09-28 TO 26-10-28
# # 28-10-28 TO 30-11-28
# # 02-12-28 TO 28-12-28
# # 30-12-28 TO 25-01-29
# # 27-01-29 TO 22-02-29
# # 24-02-29 TO 29-03-29
# # 31-03-29 TO 26-04-29
# # 28-04-29 TO 31-05-29
# # 02-06-29 TO 28-06-29
# # 30-06-29 TO 26-07-29
# # 28-07-29 TO 30-08-29
# # 01-09-29 TO 27-09-29
# # 29-09-29 TO 25-10-29
# # 27-10-29 TO 29-11-29
# # 01-12-29 TO 27-12-29
# # 29-12-29 TO 31-01-30
# # 02-02-30 TO 28-02-30
# # 02-03-30 TO 28-03-30
# # 30-03-30 TO 25-04-30
# # 27-04-30 TO 30-05-30"""
# #
# # date_format = "%d-%m-%y"
# #
# # # for s in dates.split("\n"):
# # #     start_date = datetime.datetime.strptime(s.split(" ")[0], date_format) - datetime.timedelta(days=1)
# # #     end_date = datetime.datetime.strptime(s.split(" ")[-1], date_format) - datetime.timedelta(days=1)
# # #
# # #     print(f"{start_date.strftime(date_format)} ---- {end_date.strftime(date_format)}")
# #
# # from datetime import datetime, timedelta
# #
# #
# # def get_last_friday(year, month):
# #     last_day = datetime(year, month, 1) + timedelta(days=32)
# #     last_day = last_day.replace(day=1) - timedelta(days=1)
# #     days_back = (last_day.weekday() - 4) % 7
# #     return last_day - timedelta(days=days_back)
# #
# #
# # def get_last_thursday(year, month):
# #     last_day = datetime(year, month, 1) + timedelta(days=32)
# #     last_day = last_day.replace(day=1) - timedelta(days=1)
# #     days_back = (last_day.weekday() - 3) % 7
# #     return last_day - timedelta(days=days_back)
# #
# #
# # def print_date_ranges_with_dayname(start_year, start_month, end_year, end_month):
# #     current_year, current_month = start_year, start_month
# #
# #     while (current_year < end_year) or (current_year == end_year and current_month <= end_month):
# #         start_date = get_last_friday(current_year, current_month)
# #         next_month = current_month + 1 if current_month < 12 else 1
# #         next_year = current_year if current_month < 12 else current_year + 1
# #         end_date = get_last_thursday(next_year, next_month)
# #
# #         # Get day names
# #         start_dayname = start_date.strftime('%A')  # Get day name for start date
# #         end_dayname = end_date.strftime('%A')  # Get day name for end date
# #
# #         # Print the range with day names
# #         print(f"{start_date.strftime('%d-%m-%y')} TO {end_date.strftime('%d-%m-%y')}")
# #
# #         # Advance to the next month
# #         current_month = next_month
# #         current_year = next_year
# #
# #
# # # Generate and print date ranges with day names
# # print_date_ranges_with_dayname(2024, 12, 2030, 1)
# from time import sleep
# from selenium import webdriver
# from selenium.webdriver.support.ui import WebDriverWait
# from selenium.webdriver.support import expected_conditions as ec
# from selenium.webdriver.common.by import By
# from selenium.webdriver.chrome.options import Options
# from selenium.common.exceptions import TimeoutException
#
#
# # for close and LTP filling from NSE        // make this a function that I can reuse in algo.py
# options = Options()
# options.add_argument("--disable-blink-features=AutomationControlled")
#
# # Exclude the collection of enable-automation switches
# options.add_experimental_option("excludeSwitches", ["enable-automation"])
#
# # Turn-off userAutomationExtension
# options.add_experimental_option("useAutomationExtension", False)
#
# cash_close_list = ["AARTIIND", "ADANIENT", "APOLLOTYRE", "BAJAJFINSV", "BAJFINANCE", "BANDHANBNK", "BANKBARODA", "COALINDIA",
#                    "DLF", "EICHERMOT", "FEDERALBNK", "HCLTECH", "HDFCBANK", "HINDALCO", "ICICIBANK", "INDUSINDBK", "INFY",
#                    "JINDALSTEL", "LICHSGFIN", "M%26M", "M%26MFIN", "NTPC", "RELIANCE", "SBIN", "SUNTV", "TATACHEM", "TATAMOTORS",
#                    "TATAPOWER", "TATASTEEL", "ULTRACEMCO"]
#
# # cash_close_list1 = ["M%26M"]
#
# manual = []         # list to keep track of the shares whose values selenium couldn't get
# close = []
# ltp = []
# ltp_xpath = '/html/body/div[11]/div/div/section/div/div/div/div/div/div[2]/div/section/div/div/div/aside[2]/div/div/table/tbody/tr/td[5]'
#
# print(f"Share: Close-LTP")
# for share in cash_close_list:
#     driver = webdriver.Chrome(options=options)
#
#     driver.get(f"https://www.nseindia.com/get-quotes/equity?symbol={share}")
#
#     try:
#         sleep(2)
#         myElem = WebDriverWait(driver, 20).until(ec.presence_of_element_located((By.ID, 'quoteLtp')))
#         # sleep(5)
#         close_val = driver.find_element(By.ID, "quoteLtp").text
#         ltp_val = driver.find_element(By.XPATH, ltp_xpath).text
#
#         while close_val == '' and ltp_val == '':
#             driver.refresh()
#             WebDriverWait(driver, 10).until(ec.presence_of_element_located((By.ID, 'quoteLtp')))
#             WebDriverWait(driver, 10).until(ec.presence_of_element_located((By.XPATH, ltp_xpath)))
#             close_val = driver.find_element(By.ID, "quoteLtp").text
#             ltp_val = driver.find_element(By.XPATH, ltp_xpath).text
#             sleep(0.5)
#
#         close_val = close_val.replace(",", "")
#         ltp_val = ltp_val.replace(",", "")
#
#         # truncating last 0
#         if close_val[-1] == '0':
#             close_val = close_val[:-1]
#
#         # truncating last 0
#         if ltp_val[-1] == '0':
#             ltp_val = ltp_val[:-1]
#
#         close.append(close_val)
#         ltp.append(ltp_val)
#
#         print(f'{share}: {close_val}-{ltp_val}')
#         if close_val == '':
#             manual.append(f"{share}Close")
#
#         if ltp_val == '':
#             manual.append(f"{share}LTP")
#
#     except TimeoutException:        # added temp fix for timeoutexception, will need to check if it works properly or nah
#         close.append('')
#         manual.append('')
#         manual.append(f"{share} Timeout")
#         print(f"Loading took too much time for {share}!")
#
#     driver.close()
#
# print(close)
# print(ltp)
# print(manual)
import os

algo_share_list = ['AARTIIND', 'ABB', 'ABCAPITAL', 'ABFRL', 'ADANIENT', 'ADANIPORTS', 'ALKEM', 'AMBUJACEM',
                     'APOLLOHOSP', 'APOLLOTYRE', 'ASHOKLEY', 'ASTRAL', 'ATUL', 'AUBANK', 'AUROPHARMA', 'BAJAJAUTO', 'BAJAJFINSV',
                     'BAJFINANCE', 'BALKRISIND', 'BALRAMCHIN', 'BANDHANBNK', 'BANKBARODA', 'BATAINDIA', 'BEL',
                     'BHARATFORG', 'BHEL', 'BIOCON', 'BRITANNIA', 'BSOFT', 'CANBK', 'CANFINHOME', 'CHAMBLFERT', 'CHOLAFIN',
                     'CIPLA', 'COFORGE', 'CONCOR', 'COROMANDEL', 'CROMPTON', 'CUMMINSIND', 'DABUR', 'DALBHARAT',
                     'DEEPAKFERT', 'DEEPAKNTR', 'DELTACORP', 'DIVISLAB', 'DIXON', 'DLF', 'DRREDDY', 'ESCORTS',
                     'EXIDEIND', 'GLENMARK', 'GLS', 'GNFC', 'GODREJCP', 'GODREJPROP', 'GRANULES', 'GRASIM', 'GUJGASLTD',
                     'HAL', 'HAVELLS', 'HCLTECH', 'HDFCAMC', 'HDFCLIFE', 'HINDALCO', 'HINDCOPPER', 'ICICIGI',
                     'ICICIPRULI', 'IEX', 'IGL', 'INDHOTEL', 'INDIACEM', 'INDIAMART', 'INDIGO', 'INDUSINDBK',
                     'INDUSTOWER', 'INTELLECT', 'IPCALAB', 'JINDALSTEL', 'JKCEMENT', 'JSWSTEEL', 'JUBLFOOD',
                     'KOTAKBANK', 'LALPATHLAB', 'LAURUSLABS', 'LICHSGFIN', 'LTIM', 'LTTS', 'LUPIN', 'M&MFIN',
                     'MANAPPURAM', 'MARICO', 'MCDOWELL-N', 'MCX', 'METROPOLIS', 'MFSL', 'MGL', 'MPHASIS', 'MUTHOOTFIN',
                     'NAM-INDIA', 'NAUKRI', 'NAVINFLUOR', 'NMDC', 'NTPC', 'OBEROIRLTY', 'ONGC', 'PEL', 'PERSISTENT', 'PETRONET',
                     'PIDILITIND', 'POLYCAB', 'POWERGRID', 'RAIN', 'RAMCOCEM', 'RBLBANK', 'RECLTD', 'SBICARD',
                     'SBILIFE', 'SIEMENS', 'SRF', 'STAR', 'SUNPHARMA', 'SYNGENE', 'TATACOMM', 'TATAMOTORS', 'TCS', 'TECHM',
                     'TITAN', 'TORNTPHARM', 'TORNTPOWER', 'TRENT', 'TVSMOTOR', 'UBL', 'ULTRACEMCO', 'UPL', 'VEDL', 'VOLTAS',
                     'ZEEL', 'ZYDUSLIFE']

cash_share_list = ["AARTIIND", "ADANI", "APOLLO", "BAJFINSV", "BAJFIN", "BANBK", "BARODA", "COALIND", "DLF", "EICHER",
                   "FEDBANK", "HCL", "HDFC", "HIND", "ICICI", "INDUSIND", "INFY", "JIND", "LIC", "M&M", "M&MFIN", "NTPC",
                   "REL", "SBIN", "SUNTV", "TCHEM", "TM", "TP", "TS", "ULTRA"]

c = ['449.0', '6526.85', '178.2', '282.05', '2434.1', '1148.45', '5265.0', '535.2', '6767.0', '452.1', '208.43', '1484.0', '6732.5', '604.95', '1199.9', '8532.9', '1737.0', '7427.5', '2719.9', '499.0', '152.21', '231.62', '1307.9', '285.35', '1218.05', '216.41', '401.2', '4871.1', '532.4', '100.95', '708.0', '497.35', '1246.35', '1447.95', '8651.0', '771.75', '1849.8', '362.1', '2940.35', '521.5', '1755.0', '1176.1', '2360.0', '111.08', '6011.7', '17520.5', '760.0', '1301.0', '3574.0', '390.85', '1501.55', '1100.2', '561.75', '1160.0', '2425.05', '598.85', '2398.0', '500.0', '4190.05', '1600.05', '1794.95', '4038.45', '633.6', '618.2', '250.0', '1918.0', '644.6', '172.76', '403.5', '792.0', '380.15', '2270.0', '4112.65', '970.0', '374.0', '917.2', '1549.7', '932.0', '4500.0', '922.0', '694.95', '1918.0', '2825.0', '584.75', '569.0', '5827.85', '5373.0', '2132.9', '267.45', '194.4', '657.6', '1460.5', '5989.4', '1890.0', '1075.0', '1309.0', '2802.35', '2143.0', '686.55', '7712.75', '3783.95', '67.14', '334.95', '2002.0', '268.69', '1022.7', '6075.0', '321.75', '2791.0', '6800.0', '307.0', '158.5', '884.0', '157.79', '489.35', '760.0', '1497.0', '6129.0', '2620.05', '610.5', '1773.0', '823.25', '1725.1', '773.95', '4080.0', '1675.0', '3393.0', '3174.0', '1476.0', '6079.95', '2288.0', '2060.6', '10635.0', '554.95', '460.3', '1548.4', '120.97', '996.0']
l = ['449.15', '6525.25', '178.33', '282.65', '2442.8', '1148.7', '5238.0', '535.35', '6780.85', '451.3', '209.22', '1483.45', '6729.95', '606.0', '1202.05', '8544.4', '1735.2', '7440.1', '2721.5', '502.05', '152.44', '232.12', '1302.85', '285.8', '1217.2', '216.92', '401.55', '4885.4', '533.1', '101.04', '708.0', '497.9', '1246.9', '1445.45', '8662.1', '774.25', '1857.8', '362.3', '2943.3', '521.6', '1752.05', '1175.95', '2360.0', '111.09', '6000.6', '17559.45', '760.6', '1302.35', '3575.4', '391.55', '1503.6', '1104.7', '561.2', '1166.65', '2428.15', '597.75', '2402.0', '499.9', '4187.5', '1601.1', '1796.2', '4040.4', '634.05', '618.15', '251.24', '1922.7', '641.3', '172.98', '402.4', '793.0', '379.25', '2270.0', '4114.3', '970.7', '375.6', '919.1', '1547.1', '933.8', '4503.2', '919.2', '694.55', '1920.5', '2820.5', '584.25', '569.9', '5825.3', '5373.3', '2130.65', '267.65', '193.71', '657.45', '1457.85', '6026.9', '1895.55', '1075.85', '1307.1', '2805.25', '2142.2', '689.7', '7770.3', '3777.15', '67.11', '336.05', '2002.9', '269.36', '1025.3', '6090.9', '321.2', '2787.75', '6807.55', '306.25', '158.94', '884.05', '158.18', '490.75', '761.5', '1499.7', '6136.45', '2612.1', '610.75', '1778.25', '821.7', '1729.9', '774.35', '4077.8', '1674.6', '3382.25', '3164.25', '1482.95', '6090.0', '2292.95', '2059.4', '10624.45', '552.85', '460.95', '1546.35', '120.95', '995.5']

# for i in l:
    # print(i)
cash_30_min_list = {"AARTIIND": 2, "ABB": 3, "ADANI": 3, "APOLLO": 4, "ASHOKLEY": 12, "BAJFINSV": 5, "BAJFIN": 6, "BANBK": 7,
                    "BHEL": 27, "BARODA": 8, "BN": 2, "CHAMBAL": 33, "COALIND": 9, "DIXON": 47, "DLF": 10, "EICHER": 11, "ESCORTS": 50, "FEDBANK": 12, "HCL": 13,
                    "HINDALCO": 15, "IGL": 70, "INDUSIND": 17, "JIND": 19, "LIC": 20, "M&M": 21, "M&MFIN": 22,
                    "NIFTY": 3, "NTPC": 23, "ONGC": 106, "RECLTD": 116, "REL": 24, "SBIN": 25, "SUNTV": 26, "TM": 28,
                    "TP": 29, "TS": 30, "VEDL": 136}
for i, l in enumerate(os.listdir(r'E:\Daily Data work\hourlys 30 minute CASH\2025\JAN\31.01.25')):
    print(l[:-4] == list(cash_30_min_list.keys())[i])
