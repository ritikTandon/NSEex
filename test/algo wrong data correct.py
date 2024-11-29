import openpyxl as xl
import pyautogui as pg
from time import sleep

from openpyxl.formatting.rule import CellIsRule
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException
from selenium import webdriver
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

algo_share_list = ['AARTIIND', '02 ABB', 'ABCAPITAL', 'ABFRL', 'ADANIENT', 'ADANIPORTS', 'ALKEM', 'AMBUJACEM',
                         'APOLLOHOSP', 'APOLLOTYRE', '03 ASHOKLEY', 'ASTRAL', 'ATUL', 'AUBANK', 'AUROPHARMA', 'BAJAJFINSV',
                         'BAJFINANCE', 'BALKRISIND', 'BALRAMCHIN', 'BANDHANBNK', 'BANKBARODA', 'BATAINDIA', 'BEL',
                         'BHARATFORG', '04 BHEL', 'BIOCON', 'BRITANNIA', 'BSOFT', 'CANBK', 'CANFINHOME', 'CHAMBLFERT', 'CHOLAFIN',
                         'CIPLA', 'COFORGE', 'CONCOR', 'COROMANDEL', 'CROMPTON', 'CUMMINSIND', 'DABUR', 'DALBHARAT',
                         'DEEPAKFERT', 'DEEPAKNTR', 'DELTACORP', 'DIVISLAB', '05 DIXON', 'DLF', 'DRREDDY', 'ESCORTS',
                         'EXIDEIND', 'GLENMARK', 'GLS', 'GNFC', 'GODREJCP', 'GODREJPROP', 'GRANULES', 'GRASIM', 'GUJGASLTD',
                         'HAL', 'HAVELLS', 'HCLTECH', 'HDFCAMC', 'HDFCLIFE', 'HINDALCO', 'HINDCOPPER', 'ICICIGI',
                         'ICICIPRULI', 'IEX', 'IGL', 'INDHOTEL', 'INDIACEM', 'INDIAMART', 'INDIGO', 'INDUSINDBK',
                         'INDUSTOWER', 'INTELLECT', 'IPCALAB', 'JINDALSTEL', 'JKCEMENT', 'JSWSTEEL', 'JUBLFOOD',
                         'KOTAKBANK', 'LALPATHLAB', 'LAURUSLABS', 'LICHSGFIN', 'LTIM', 'LTTS', 'LUPIN', 'M%26MFIN',
                         'MANAPPURAM', 'MARICO', 'MCDOWELL-N', 'MCX', 'METROPOLIS', 'MFSL', 'MGL', 'MPHASIS', 'MUTHOOTFIN',
                         'NAM-INDIA', 'NAUKRI', 'NAVINFLUOR', 'NMDC', 'NTPC', 'OBEROIRLTY', '09 ONGC', 'PEL', 'PERSISTENT', 'PETRONET',
                         'PIDILITIND', 'POLYCAB', 'POWERGRID', 'RAIN', 'RAMCOCEM', 'RBLBANK', '10 RECLTD', 'SBICARD',
                         'SBILIFE', 'SIEMENS', 'SRF', 'STAR', 'SUNPHARMA', 'SYNGENE', 'TATACOMM', 'TATAMOTORS', 'TECHM',
                         'TORNTPHARM', 'TORNTPOWER', 'TRENT', 'TVSMOTOR', 'UBL', 'ULTRACEMCO', 'UPL', 'VEDL', 'VOLTAS',
                         'ZEEL', 'ZYDUSLIFE']

daily_cell_range = "H3:K3000"
weekly_cell_range = "E4:I1000"
other_cell_range = "F4:J1000"

font = Font(name='Arial', size=11, bold=True)  # text font
heading_font = Font(name='Arial', size=11, bold=True, color="FFFFFF")  # heading
fill = PatternFill(patternType='solid', fgColor="0000ff")  # blue
red = Font("Arial", 11, color='ff0000', bold=True)
blue = Font("Arial", 11, color="0000ff", bold=True)
alignment = Alignment(horizontal='center')
border = Border(
    left=Side(style='thin'),
    right=Side(style='thin'),
    top=Side(style='thin'),
    bottom=Side(style='thin')
)

red_text_rule = CellIsRule(operator='lessThan', formula=['0'], font=red)
blue_text_rule = CellIsRule(operator='greaterThanOrEqual', formula=['0'], font=blue)

sh_row = 2
d_row = 1101
w_row = 221
m_row = 54
cl_row = 53
w_range = [1097, 1101]
m_range = [1091, 1110]
cl_range = [1091, 1110]
sh_count = 0

xp_high = '/html/body/div[11]/div/div/section/div/div/div/div/div/div[2]/div/div/div/div[2]/div/div[1]/section/div/div[3]/div/table/tbody/tr[2]/td[4]'
xp_low = '/html/body/div[11]/div/div/section/div/div/div/div/div/div[2]/div/div/div/div[2]/div/div[1]/section/div/div[3]/div/table/tbody/tr[2]/td[5]'

options = Options()
# options.add_argument('--headless=new')
options.add_argument("--disable-blink-features=AutomationControlled")
options.add_argument("--start-maximized")

# Exclude the collection of enable-automation switches
options.add_experimental_option("excludeSwitches", ["enable-automation"])

# Turn-off userAutomationExtension
options.add_experimental_option("useAutomationExtension", False)

algo_close_list = ['AARTIIND', 'ABB', 'ABCAPITAL', 'ABFRL', 'ADANIENT', 'ADANIPORTS', 'ALKEM', 'AMBUJACEM',
                     'APOLLOHOSP', 'APOLLOTYRE', 'ASHOKLEY', 'ASTRAL', 'ATUL', 'AUBANK', 'AUROPHARMA', 'BAJAJ-AUTO', 'BAJAJFINSV',
                     'BAJFINANCE', 'BALKRISIND', 'BALRAMCHIN', 'BANDHANBNK', 'BANKBARODA', 'BATAINDIA', 'BEL',
                     'BHARATFORG', 'BHEL', 'BIOCON', 'BRITANNIA', 'BSOFT', 'CANBK', 'CANFINHOME', 'CHAMBLFERT', 'CHOLAFIN',
                     'CIPLA', 'COFORGE', 'CONCOR', 'COROMANDEL', 'CROMPTON', 'CUMMINSIND', 'DABUR', 'DALBHARAT',
                     'DEEPAKFERT', 'DEEPAKNTR', 'DELTACORP', 'DIVISLAB', 'DIXON', 'DLF', 'DRREDDY', 'ESCORTS',
                     'EXIDEIND', 'GLENMARK', 'GLS', 'GNFC', 'GODREJCP', 'GODREJPROP', 'GRANULES', 'GRASIM', 'GUJGASLTD',
                     'HAL', 'HAVELLS', 'HCLTECH', 'HDFCAMC', 'HDFCLIFE', 'HINDALCO', 'HINDCOPPER', 'ICICIGI',
                     'ICICIPRULI', 'IEX', 'IGL', 'INDHOTEL', 'INDIACEM', 'INDIAMART', 'INDIGO', 'INDUSINDBK',
                     'INDUSTOWER', 'INTELLECT', 'IPCALAB', 'JINDALSTEL', 'JKCEMENT', 'JSWSTEEL', 'JUBLFOOD',
                     'KOTAKBANK', 'LALPATHLAB', 'LAURUSLABS', 'LICHSGFIN', 'LTIM', 'LTTS', 'LUPIN', 'M%26MFIN',
                     'MANAPPURAM', 'MARICO', 'UNITDSPR', 'MCX', 'METROPOLIS', 'MFSL', 'MGL', 'MPHASIS', 'MUTHOOTFIN',
                     'NAM-INDIA', 'NAUKRI', 'NAVINFLUOR', 'NMDC', 'NTPC', 'OBEROIRLTY', 'ONGC', 'PEL', 'PERSISTENT', 'PETRONET',
                     'PIDILITIND', 'POLYCAB', 'POWERGRID', 'RAIN', 'RAMCOCEM', 'RBLBANK', 'RECLTD', 'SBICARD',
                     'SBILIFE', 'SIEMENS', 'SRF', 'STAR', 'SUNPHARMA', 'SYNGENE', 'TATACOMM', 'TATAMOTORS', 'TCS', 'TECHM',
                     'TITAN', 'TORNTPHARM', 'TORNTPOWER', 'TRENT', 'TVSMOTOR', 'UBL', 'ULTRACEMCO', 'UPL', 'VEDL', 'VOLTAS',
                     'ZEEL', 'ZYDUSLIFE']


doub = ['ADANIPORTS', 'CHOLAFIN']
add = 0
sh_row = 34
for share in algo_close_list:
    # if share takes 2 line as name
    if share in doub:
        add = 20
    left = (1105, 530+add)
    ll = (888, 575+add)  # 6
    ldate = (1052, 701+add)
    hist_data = (593, 439+add)
    right = (1387, 532+add)
    rl = (1172, 573+add)  # 7
    rdate = (1334, 702+add)
    flter = (1467, 527+add)
    hlwb = xl.load_workbook(rf'C:\Users\admin\PycharmProjects\daily data\test\hl_test.xlsx')
    hlsh = hlwb['Sheet2']
    driver = webdriver.Chrome(options=options)

    driver.get(f"https://www.nseindia.com/get-quotes/equity?symbol={share}")



    try:
        myElem = WebDriverWait(driver, 20).until(ec.presence_of_element_located((By.ID, 'historic_data')))
        sleep(2)
        pg.click(hist_data)
        sleep(2)
        # pg.click(left)
        # sleep(2)
        # pg.click(ll, clicks=6, interval=0.6)
        # sleep(2)
        # pg.click(ldate)
        # sleep(2)
        # pg.click(right)
        # sleep(2)
        # pg.click(rl, clicks=7, interval=0.6)
        # sleep(2)
        # pg.click(rdate)
        # sleep(2)
        # pg.click(flter)
        # sleep(2)
        he = WebDriverWait(driver, 20).until(ec.presence_of_element_located((By.XPATH, xp_high)))
        high = float(driver.find_element(By.XPATH, xp_high).text.replace(",", ""))
        low = float(driver.find_element(By.XPATH, xp_low).text.replace(",", ""))

        hlsh.cell(sh_row, 2).value = high
        hlsh.cell(sh_row, 3).value = low
        sh_row += 1

        hlwb.save(rf'C:\Users\admin\PycharmProjects\daily data\test\hl_test.xlsx')
        print(f"{high} {low} {share}")

    except TimeoutException:
        print(f"Loading took too much time for {share}!")

    driver.close()


# # copying high and low from algo hl sheet (after putting correct high and low values for that date)
# for share in algo_share_list:
#     hlwb = xl.load_workbook(rf'C:\Users\admin\PycharmProjects\daily data\test\hl_test.xlsx')
#     hlsh = hlwb['Sheet2']
#
#     wb = xl.load_workbook(rf'E:\Daily Data work\ALGORITHM\ALGORITHM OLD\{share}.xlsx')
#     sh = wb['D']
#
#     sh.cell(d_row, 2).value = hlsh.cell(sh_row, 2).value        # high copy
#     sh.cell(d_row, 3).value = hlsh.cell(sh_row, 3).value        # low copy
#
#     wb.save(rf'C:\Users\admin\PycharmProjects\daily data\test\algo\{share}.xlsx')
#     sh_row += 1
#     sh_count += 1
#
#     print(f"{share} done")
#
# print(f"{sh_count} shares done")