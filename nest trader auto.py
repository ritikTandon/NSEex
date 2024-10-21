import shutil
import pyautogui as pg
import pyperclip as pc
from date_variables import date, mnth, yr
from time import sleep
import time
import numpy as np

# timing the start of execution
begin = time.time()


# ------------------------------------------- functions for code readability--------------------------------------------
# ctrl+shift+h to open vwap stats
def vwap():
    pg.hotkey('ctrl', 'shift', 'h')
    sleep(1)


def write(name):  # copying and writing the text
    pc.copy(name)
    pg.hotkey("ctrl", "v")


# function to check if the screen has changed, indicating that the vwap statistics have loaded in, and we need not wait longer
def check_change():
    changed = False  # bool representing if screen has changed from black to showing vwap statistics

    # waiting for screen to change before saving
    while not changed:
        pixel = pg.screenshot(
            region=(imp_coord_dict["color change"][0], imp_coord_dict["color change"][1], 6, 8))  # taking pixel average

        pixels = np.array(pixel)

        # Calculate the average color
        average_color = pixels.mean(axis=(0, 1)).astype(int)

        # if color is green (screen has updated)
        if 200 > average_color[1] > 10:
            print("Displayed the vwap data")
            changed = True

    # opening 'save as window'
    sleep(1)
    pg.click(imp_coord_dict["color change"])
    pg.hotkey('shift', 'e')  # write to excel shortcut
    sleep(1)

    # showing that we are outside the vwap statistics after we saved the data
    print("Outside")


# function to close the excel file that just opened after saving the vwap data
def excel_close():
    sleep(2)
    pg.rightClick(imp_coord_dict['excel icon taskbar'])
    sleep(1)
    pg.click(imp_coord_dict['excel close'])
    sleep(1)


# saving the file
def save_share(name, pth=None):
    # if path is mentioned, save it at that particular path. Until changed, all shares will default to this path because of how the trading app works
    if pth:
        print(name)
        pg.click(imp_coord_dict["path"])
        write(pth)
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


# saving the daily consolidated sheets
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
    sleep(1)
    pg.press('enter')
    sleep(1)
    pg.click(imp_coord_dict["share name textbox"])
    sleep(1)
    write(save_name)
    sleep(1)
    pg.click(imp_coord_dict['save'])
    sleep(1)
    pg.hotkey("esc")
    pg.hotkey("esc")
    sleep(1)


# saving all the daily sheets
def sheetSave():
    save_daily_sheet("csh")
    save_daily_sheet("fo1")
    save_daily_sheet("algo")
    excel_close()


# DO NOT TOUCH THE 'SAVE AS' WINDOW IN NEST AND IF YOU DO, CHANGE THE 'path' AND 'share name' COORDINATES ACCORDINGLY

imp_coord_dict = {"nest icon taskbar": (797, 1058), "EQ": (152, 992), "first share": (73, 173), "FO": (184, 992),
                  "FO1": (213, 992), "30minFO": (266, 992), "30minCash": (33, 992), "time interval": (537, 58),
                  "get stats": (629, 55), "color change": (929, 212), "path": (1298, 312),
                  "share name textbox": (831, 699),
                  "save": (1415, 774), "excel icon taskbar": (901, 1054), "excel close": (897, 1014), "TL": (635, 265),
                  "BR": (1579, 796), "ALGO": (311, 993)}    # TL and BR are "save as" window cords

currentMouseX, currentMouseY = pg.position()  # Returns two integers, the x and y of the mouse cursor's current position.

EQ_shares = ["AARTIIND", "ADANI", "APOLLO", "BAJFINSV", "BAJFIN", "BANBK", "BARODA", "COALIND", "DLF", "EICHER", "FEDBANK",
             "HCL", "HDFC", "HIND", "ICICI", "INDUSIND", "INFY", "JIND", "LIC", "M&M", "M&MFIN", "NTPC", "REL", "SBIN",
             "SUNTV", "TCHEM", "TM", "TP", "TS", "ULTRA"]

FO_shares = ["ADANI", "APORT", "APOLLO", "AURO", "AXIS", "BAJAJ", "BARODA", "BN", "AIRTEL", "BHEL", "CANBK", "COALIND",
             "DLF", "DRREDDY", "EICHER", "HCL", "HDFC", "HIND", "HINDUNLVR", "ICICI", "INDUSIND", "JIND", "NIFTY", "REL",
             "SBIN", "TCHEM", "TCON", "TM", "TS", "TCS", "TITAN", "ULTRA", "VEDL"]

EQ_30_min_shares = ["AARTIIND", "ABB", "ADANI", "APOLLO", "ASHOKLEY", "BAJFINSV", "BAJFIN", "BANBK", "BARODA", "BN",
                    "BHEL", "DIXON", "DLF", "EICHER", "ESCORTS", "FEDBANK", "HCL", "HINDALCO", "IGL", "INDUSIND", "JIND",
                    "LIC", "M&M", "M&MFIN", "NIFTY", "NTPC", "ONGC", "RECLTD", "SBIN", "SUNTV", "TM", "TP", "TS", "VEDL"]

# same as algo share list in algo.py
ALGO_1_min_shares = ['AARTIIND', 'ABB', 'ABCAPITAL', 'ABFRL', 'ADANIENT', 'ADANIPORTS', 'ALKEM', 'AMBUJACEM',
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
sheetSave()

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
        sleep(1)
        pg.press('enter')
        sleep(2)

        check_change()
        save_share(share, path)
        first = False
        sleep(1)

    else:
        print('after first')
        pg.press('enter')
        sleep(2)
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
        sleep(2)

        check_change()
        save_share(share, path)
        first = False
        sleep(1)

    else:
        print('after first')
        pg.press('enter')
        sleep(2)
        check_change()
        save_share(share)
        sleep(1)

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
        sleep(1)
        pg.press('enter')
        sleep(2)

        check_change()
        save_share(share, path)
        first = False
        sleep(1)

    else:
        print('after first')
        pg.press('enter')
        sleep(2)
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
        sleep(1)
        pg.press('enter')

        sleep(1)
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

# timing the end of execution
end = time.time()
print(f"Total runtime of the program is {int((end - begin) // 60)} minutes and {int((end - begin) % 60)} seconds")
