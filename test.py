# import openpyxl as xl
#
# wb = xl.load_workbook(r'C:\Users\admin\PycharmProjects\daily data\test.xlsx')
# sheet = wb["Sheet1"]
#
# cell = sheet.cell(1,1)
# cell.number_format = 'hh:mm AM/PM'
#
# print(cell.value)
# print(cell.number_format)
#
# wb.save(r'C:\Users\admin\PycharmProjects\daily data\test.xlsx')
# wb.close()
#
# wb = xl.load_workbook(r'C:\Users\admin\PycharmProjects\daily data\test.xlsx')
# sheet = wb["Sheet1"]
#
# cell = sheet.cell(1,1)
# print(cell.value)

# from time import sleep
# from selenium import webdriver
# from selenium.webdriver.support.ui import WebDriverWait
# from selenium.webdriver.support import expected_conditions as EC
# from selenium.webdriver.common.by import By
# from selenium.webdriver.chrome.options import Options
# from selenium.common.exceptions import TimeoutException
#
#
# options = Options()
# options.add_argument('--headless=new')
#
# cash_close_list = ["ADANIENT", "APOLLOTYRE", "BAJAJFINSV", "BAJFINANCE", "BANDHANBNK", "BANKBARODA", "COALINDIA", "DLF",
#                    "EICHERMOT", "FEDERALBNK", "HCLTECH", "HDFCBANK", "HINDALCO", "ICICIBANK", "INDUSINDBK", "INFY",
#                    "JINDALSTEL", "LICHSGFIN", "M%26M", "M%26M", "RELIANCE", "SBIN", "SUNTV", "TATACHEM", "TATAMOTORS",
#                    "TATAPOWER", "TATASTEEL", "ULTRACEMCO"]
#
# cash_close_list1 = ["M%26M"]
#
# ltp = []
#
# for share in cash_close_list:
#     driver = webdriver.Chrome(options=options)
#
#     driver.get(f"https://www.nseindia.com/get-quotes/equity?symbol={share}")
#
#     try:
#         sleep(2)
#         myElem = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'quoteLtp')))
#         sleep(5)
#         ltp_val = driver.find_element(By.ID, "quoteLtp").text
#         ltp.append(ltp_val)
#
#         print(f'{share}: {ltp_val}')
#
#     except TimeoutException:
#         print("Loading took too much time!")
#
#     driver.close()
#
# print(ltp)

# cash_share_list = ["ADANI", "APOLLO", "BAJFINSV", "BAJFIN", "BANBK", "BARODA", "COALIND", "DLF", "EICHER", "FEDBANK",
#                    "HCL", "HDFC", "HIND", "ICICI", "INDUSIND", "INFY", "JIND", "LIC", "M&M", "M&MFIN", "REL", "SBIN",
#                    "SUNTV", "TCHEM", "TM", "TP", "TS", "ULTRA"]
#
# for share in cash_share_list:
#     print(share)

# def isValid(s: str) -> bool:
#     stack = []
#     bol = False
#
#     if len(s) == 1:
#         return bol
#
#     for i in s:
#         stack.append(i)
#         top = len(stack) - 1
#
#         if len(stack) == 1 and i in [')', '}', ']']:
#             return False
#
#         if len(stack) > 1:
#             if i in [')', '}', ']']:
#                 if i == ')' and stack[top-1] == "(":
#                     stack.pop()
#                     stack.pop()
#                     bol = True
#                 elif i == '}' and stack[top-1] == "{":
#                     stack.pop()
#                     stack.pop()
#                     bol = True
#                 elif i == ']' and stack[top-1] == "[":
#                     stack.pop()
#                     stack.pop()
#                     bol = True
#
#                 else:
#                     return False
#
#     if len(stack) > 0:
#         return False
#
#     return bol
#
#
# print(isValid("([]){"))

# s = "1234"
#
# print(s[len(s)-1:len(s)])
# print(s[:len(s)-1])
#
# share_dict = {"BANKNIFTY": 4, "NIFTY": 10, "ADANIENT": 2, "AUROPHARMA": 3, "CANBK": 5, "DLF": 6, "HINDALCO": 7,
#               "ICICIBANK": 8, "JINDALSTEL": 9, "RELIANCE": 11, "SBIN": 12, "TATACONSUM": 13, "TATAMOTORS": 14,
#               "TATASTEEL": 15, "TCS": 16, "TITAN": 17}
#
# print(share_dict.keys()[0])