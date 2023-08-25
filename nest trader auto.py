import shutil
import pyautogui as pg
import pyperclip as pc
from date_variables import date, mnth, yr
from time import sleep


# functions for code readability
# ctrl+shift+h to open vwap stats
def vwap():
    pg.hotkey('ctrl', 'shift', 'h')
    sleep(1)


def write(name):        # copying and writing the text
    pc.copy(name)
    pg.hotkey("ctrl", "v")


def check_change():
    changed = False     # bool representing if screen has changed from black to showing vwap statistics

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

    sleep(2)
    pg.rightClick(imp_coord_dict['excel icon taskbar'])
    sleep(1)
    pg.click(imp_coord_dict['excel close'])
    sleep(1)
    # closing out of vwap statistics and moving to next share
    pg.press('esc', presses=4, interval=0.5)
    pg.press('down')
    sleep(1)


# DO NOT TOUCH THE 'SAVE AS' WINDOW IN NEST AND IF YOU DO, CHANGE THE 'path' AND 'share name' COORDINATES ACCORDINGLY


imp_coord_dict = {"nest icon taskbar": (787, 1079), "EQ": (152, 992), "first share": (73, 173), "FO": (184, 992),
                  "FO1": (213, 992), "30minFO": (266, 992), "30minCash": (33, 992), "time interval": (537, 58),
                  "get stats": (629, 55), "color change": (935, 295), "path": (647, 547), "share name textbox": (297, 931),
                  "save": (789, 1003), "excel icon taskbar": (901, 1054), "excel close": (897, 1014), "": (), "": (), "": (), "": (), }

currentMouseX, currentMouseY = pg.position()  # Returns two integers, the x and y of the mouse cursor's current position.

EQ_shares = ["ADANI", "APOLLO", "BAJFINSV", "BAJFIN", "BANBK", "BARODA", "COALIND", "DLF", "EICHER", "FEDBANK",
             "HCL", "HDFC", "HIND", "ICICI", "INDUSIND", "INFY", "JIND", "LIC", "M&M", "M&MFIN", "REL", "SBIN",
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

PATHS_DICT = {"EQ": rf'E:\Daily Data work\hourlys 1 minute CASH\{yr}\{mnth}\{date}',
              "FO": rf'E:\Daily Data work\hourlys 1 minute FO\{yr}\{mnth}\{date}',
              "30minCash": rf'E:\Daily Data work\hourlys 30 minute CASH\{yr}\{mnth}\{date}',
              "30minFO": rf'E:\Daily Data work\hourlys 30 minute FO\{yr}\{mnth}\{date}',
              "daily data work": rf'E:\Daily Data work'}


# main loop
# clicking on NEST Trader
pg.click(imp_coord_dict["nest icon taskbar"])
pg.press('esc')
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
        path = PATHS_DICT["EQ"]         # setting path for first share

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
        path = PATHS_DICT["FO"]         # setting path for first share

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
        path = PATHS_DICT["30minFO"]         # setting path for first share

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
        path = PATHS_DICT["30minCash"]         # setting path for first share

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
shutil.copy(rf"E:\Daily Data work\hourlys 30 minute FO\{yr}\{mnth}\{date}\NIFTY.xls", rf"E:\Daily Data work\hourlys 30 minute CASH\{yr}\{mnth}\{date}")
shutil.copy(rf"E:\Daily Data work\hourlys 30 minute FO\{yr}\{mnth}\{date}\BN.xls", rf"E:\Daily Data work\hourlys 30 minute CASH\{yr}\{mnth}\{date}")

