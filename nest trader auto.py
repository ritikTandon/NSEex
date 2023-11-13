import shutil
import pyautogui as pg
import pyperclip as pc
from date_variables import date, mnth, yr
from time import sleep
import time

begin = time.time()


# functions for code readability
# ctrl+shift+h to open vwap stats
def vwap():
    pg.hotkey('ctrl', 'shift', 'h')
    sleep(1)


def write(name):  # copying and writing the text
    pc.copy(name)
    pg.hotkey("ctrl", "v")


def check_change():
    changed = False  # bool representing if screen has changed from black to showing vwap statistics

    # waiting for screen to change before saving
    while not changed:
        PIXEL = pg.screenshot(
            region=(imp_coord_dict["color change"][0], imp_coord_dict["color change"][1], 1, 1))  # taking pixel
        color = PIXEL.getcolors()

        if color[0][1] == (0, 255, 0):  # if color is green (screen has updated)
            print("Displayed the vwap data")
            changed = True

    # opening 'save as window'
    sleep(0.5)
    pg.click(imp_coord_dict["color change"])
    pg.hotkey('shift', 'e')  # write to excel shortcut
    sleep(1)

    print("Outside")


def excel_close():
    sleep(2)
    pg.rightClick(imp_coord_dict['excel icon taskbar'])
    sleep(1)
    pg.click(imp_coord_dict['excel close'])
    sleep(1)


# saving the file
def save_share(name, path=None):
    if path:
        print(name)
        pg.click(imp_coord_dict["path"])
        write(path)
        pg.press('enter')

    sleep(0.7)
    pg.click(imp_coord_dict["share name textbox"])
    write(name)
    pg.click(imp_coord_dict['save'])

    excel_close()

    # closing out of vwap statistics and moving to next share
    pg.press('esc', presses=4, interval=0.5)
    pg.press('down')
    sleep(1)


def save_daily_sheet(save_name):
    if save_name == "csh":
        pg.click(imp_coord_dict["EQ"])
    elif save_name == "csh30":
        pg.click(imp_coord_dict["30minCash"])
    elif save_name == "fo1":
        pg.click(imp_coord_dict["FO1"])
    elif save_name == "algo":
        pg.click(imp_coord_dict["ALGO"])

    sleep(1)
    pg.click(imp_coord_dict["first share"])
    sleep(1)

    # saving
    pg.hotkey('shift', "e")
    sleep(1)
    pg.click(imp_coord_dict["path"])
    sleep(1)
    write(PATHS_DICT["daily data work"])
    pg.click(imp_coord_dict["share name textbox"])
    write(save_name)
    pg.click(imp_coord_dict['save'])
    sleep(1)
    pg.hotkey("esc")
    pg.hotkey("esc")
    sleep(1)


# DO NOT TOUCH THE 'SAVE AS' WINDOW IN NEST AND IF YOU DO, CHANGE THE 'path' AND 'share name' COORDINATES ACCORDINGLY


imp_coord_dict = {"nest icon taskbar": (787, 1079), "EQ": (152, 992), "first share": (73, 173), "FO": (184, 992),
                  "FO1": (213, 992), "30minFO": (266, 992), "30minCash": (33, 992), "time interval": (537, 58),
                  "get stats": (629, 55), "color change": (935, 295), "path": (647, 547),
                  "share name textbox": (297, 931),
                  "save": (789, 1003), "excel icon taskbar": (901, 1054), "excel close": (897, 1014), "TL": (18, 495),
                  "BR": (962, 1026), "ALGO": (311, 993)}

currentMouseX, currentMouseY = pg.position()  # Returns two integers, the x and y of the mouse cursor's current position.

EQ_shares = ["ADANI", "APOLLO", "BAJFINSV", "BAJFIN", "BANBK", "BARODA", "COALIND", "DLF", "EICHER", "FEDBANK",
             "HCL", "HDFC", "HIND", "ICICI", "INDUSIND", "INFY", "JIND", "LIC", "M&M", "M&MFIN", "NTPC", "REL", "SBIN",
             "SUNTV", "TCHEM", "TM", "TP", "TS", "ULTRA"]
# EQ_shares = ["ADANI", "APOLLO", "BAJFINSV", "BAJFIN"]

FO_shares = ["ADANI", "APORT", "APOLLO", "AURO", "AXIS", "BAJAJ", "BARODA", "BN", "AIRTEL", "BHEL", "CANBK", "COALIND",
             "DLF", "DRREDDY", "EICHER", "HCL", "HDFC", "HIND", "IBUL", "ICICI", "INDUSIND", "JIND", "NIFTY", "REL",
             "SBIN", "TCHEM", "TCON", "TM", "TS", "TCS", "TITAN", "ULTRA", "VEDL"]

FO_30_min_shares = ['APOLLO', 'BAJFINSV', 'BAJFIN', 'BARODA', 'BN', 'COALIND', 'DLF', 'EICHER', 'FEDBANK', 'HCL',
                    'HDFC',
                    'ICICI', 'INDUSIND', 'INFY', 'JIND', 'M&M', 'M&MFIN', 'NIFTY', 'REL', 'SBIN', 'SUNTV', 'TCON', 'TM',
                    'TP', 'TS', 'TITAN', 'ULTRA', 'VEDL']

EQ_30_min_shares = ["ADANI", "APOLLO", "BAJFINSV", "BAJFIN", "BANBK", "BARODA", "DLF", "EICHER",
                    "FEDBANK",
                    "HCL", "INDUSIND", "JIND", "LIC", "M&M", "M&MFIN", "SBIN", "SUNTV", "TM", "TP", "TS"]

ALGO_1_min_shares = ['AARTIIND', 'ABB', 'ABCAPITAL', 'ABFRL', 'ADANIENT', 'ADANIPORTS', 'ALKEM', 'AMBUJACEM',
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

PATHS_DICT = {"EQ": rf'E:\Daily Data work\hourlys 1 minute CASH\{yr}\{mnth}\{date}',
              "FO": rf'E:\Daily Data work\hourlys 1 minute FO\{yr}\{mnth}\{date}',
              "ALGO": rf'E:\Daily Data work\hourlys 1 minute ALGO\{yr}\{mnth}\{date}',
              "30minCash": rf'E:\Daily Data work\hourlys 30 minute CASH\{yr}\{mnth}\{date}',
              "30minFO": rf'E:\Daily Data work\hourlys 30 minute FO\{yr}\{mnth}\{date}',
              "daily data work": rf'E:\Daily Data work'}

# main loop
# clicking on NEST Trader
pg.click(imp_coord_dict["nest icon taskbar"])
pg.press('esc')
sleep(1)

# saving consolidated sheets
save_daily_sheet("csh")
save_daily_sheet("fo1")
save_daily_sheet("algo")

excel_close()

# ALGO 1 min
pg.click(imp_coord_dict["ALGO"])
sleep(1)
pg.click(imp_coord_dict["first share"])
sleep(1)

first = True
for share in ALGO_1_min_shares:
    print(share)
    vwap()
    sleep(1)

    # if it's the first share, follow the first share protocol
    if first:
        path = PATHS_DICT["ALGO"]  # setting path for first share

        pg.doubleClick(imp_coord_dict["time interval"])
        pg.press("1")
        pg.press('enter')

        check_change()
        save_share(share, path)
        first = False
        sleep(1)

    else:
        print('after first')
        pg.press('enter')
        check_change()
        save_share(share)
        sleep(1)

# EQ 1 min
pg.click(imp_coord_dict["EQ"])
sleep(1)
pg.click(imp_coord_dict["first share"])
sleep(1)

first = True
for share in EQ_shares:
    print(share)
    vwap()
    sleep(1)

    # if it's the first share, follow the first share protocol
    if first:
        path = PATHS_DICT["EQ"]  # setting path for first share

        pg.doubleClick(imp_coord_dict["time interval"])
        pg.press("1")
        pg.press('enter')

        check_change()
        save_share(share, path)
        first = False
        sleep(1)

    else:
        print('after first')
        pg.press('enter')
        check_change()
        save_share(share)
        sleep(1)

# FO 1 min
pg.click(imp_coord_dict["FO"])
sleep(1)
pg.click(imp_coord_dict["first share"])
sleep(1)

first = True
for share in FO_shares:
    print(share)
    vwap()
    sleep(1)

    # if it's the first share, follow the first share protocol
    if first:
        path = PATHS_DICT["FO"]  # setting path for first share

        pg.doubleClick(imp_coord_dict["time interval"])
        pg.press("1")
        pg.press('enter')

        check_change()
        save_share(share, path)
        first = False
        sleep(1)

    else:
        print('after first')
        pg.press('enter')
        check_change()
        save_share(share)
        sleep(1)

# FO 30 min
pg.click(imp_coord_dict["30minFO"])
sleep(1)
pg.click(imp_coord_dict["first share"])
sleep(1)

first = True
for share in FO_30_min_shares:
    print(share)
    vwap()
    sleep(1)

    # if it's the first share, follow the first share protocol
    if first:
        path = PATHS_DICT["30minFO"]  # setting path for first share

        pg.doubleClick(imp_coord_dict["time interval"])
        write("30")
        pg.press('enter')

        check_change()
        save_share(share, path)
        first = False
        sleep(1)

    else:
        print('after first')
        pg.press('enter')
        check_change()
        save_share(share)
        sleep(1)

# EQ 30 min
pg.click(imp_coord_dict["30minCash"])
sleep(1)
pg.click(imp_coord_dict["first share"])
sleep(1)

first = True
for share in EQ_30_min_shares:
    print(share)
    vwap()
    sleep(1)

    # if it's the first share, follow the first share protocol
    if first:
        path = PATHS_DICT["30minCash"]  # setting path for first share

        pg.doubleClick(imp_coord_dict["time interval"])
        write("30")
        pg.press('enter')

        check_change()
        save_share(share, path)
        first = False
        sleep(1)

    else:
        print('after first')
        pg.press('enter')
        check_change()
        save_share(share)
        sleep(1)

# copying NIFTY and BN from 30 min FO to 30 minute cash
shutil.copy(rf"E:\Daily Data work\hourlys 30 minute FO\{yr}\{mnth}\{date}\NIFTY.xls",
            rf"E:\Daily Data work\hourlys 30 minute CASH\{yr}\{mnth}\{date}")
shutil.copy(rf"E:\Daily Data work\hourlys 30 minute FO\{yr}\{mnth}\{date}\BN.xls",
            rf"E:\Daily Data work\hourlys 30 minute CASH\{yr}\{mnth}\{date}")

end = time.time()
print(f"Total runtime of the program is {(end - begin) // 60} minutes and {(end - begin) % 60} seconds")
