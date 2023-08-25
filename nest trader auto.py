import os
import openpyxl as xl
import pyautogui as pg
from date_variables import date, mnth, yr
from time import sleep

# functions for code readability
# ctrl+shift+h to open vwap stats
def vwap():
    pg.press('ctrl')
    pg.press('shift')
    pg.press('h')
    pg.press('enter')

# DO NOT TOUCH THE 'SAVE AS' WINDOW IN NEST AND IF YOU DO, CHANGE THE 'path' AND 'share name' COORDINATES ACCORDINGLY

imp_coord_dict = {"nest icon taskbar": (787, 1079), "EQ": (152, 992), "first share": (73, 173), "FO": (184, 992),
                  "FO1": (213, 992), "30minFO": (266, 992), "30minCash": (33, 992), "time interval": (537, 58),
                  "get stats": (629, 55), "color change": (935, 295), "path": (442, 546), "share name": (297, 931),
                  "save": (789, 1003), "": (), "": (), "": (), "": (), "": (), "": (),}

currentMouseX, currentMouseY = pg.position()  # Returns two integers, the x and y of the mouse cursor's current position.


EQ_shares = ["ADANI", "APOLLO", "BAJFINSV", "BAJFIN", "BANBK", "BARODA", "COALIND", "DLF", "EICHER", "FEDBANK",
                   "HCL", "HDFC", "HIND", "ICICI", "INDUSIND", "INFY", "JIND", "LIC", "M&M", "M&MFIN", "REL", "SBIN",
                   "SUNTV", "TCHEM", "TM", "TP", "TS", "ULTRA"]

FO_shares = ["ADANI", "APORT", "APOLLO", "AURO", "AXIS", "BAJAJ", "BARODA", "BN", "AIRTEL", "BHEL", "CANBK", "COALIND",
                 "DLF", "DRREDDY", "EICHER", "HCL", "HDFC", "HIND", "IBUL", "ICICI", "INDUSIND", "JIND", "NIFTY", "REL",
                 "SBIN", "TCHEM", "TCON", "TM", "TS", "TCS", "TITAN", "ULTRA", "VEDL"]


FO_30_min_shares = ['APOLLO', 'BAJFINSV', 'BAJFIN', 'BARODA', 'BN', 'COALIND', 'DLF', 'EICHER', 'FEDBANK', 'HCL', 'HDFC',
                    'ICICI', 'INDUSIND', 'INFY', 'JIND', 'M&M', 'M&MFIN', 'NIFTY', 'REL', 'SBIN', 'SUNTV', 'TCON', 'TM',
                    'TP', 'TS', 'TITAN', 'ULTRA', 'VEDL']

EQ_30_min_shares = ["BN", "NIFTY", "ADANI", "APOLLO", "BAJFINSV", "BAJFIN", "BANBK", "BARODA", "DLF", "EICHER", "FEDBANK",
                    "HCL", "INDUSIND", "JIND", "LIC", "M&M", "M&MFIN", "SBIN", "SUNTV", "TM", "TP", "TS"]

PATHS_DICT = {"EQ": rf'E:\Daily Data work\hourlys 1 minute CASH\{yr}\{mnth}\{date}\\',
              "FO": rf'E:\Daily Data work\hourlys 1 minute FO\{yr}\{mnth}\{date}\\',
              "30minCash": rf'E:\Daily Data work\hourlys 30 minute CASH\{yr}\{mnth}\{date}\\',
              "30minFO": rf'E:\Daily Data work\hourlys 30 minute FO\{yr}\{mnth}\{date}\\',
              "daily data work": rf'E:\Daily Data work'}        # do no append anything to daily data work

# path = PATHS_DICT["EQ"] + rf'{share}.xlsx'

# main loop
# clicking on NEST Trader
pg.click(imp_coord_dict["nest icon taskbar"])
sleep(1)

# EQ
pg.click(imp_coord_dict["EQ"])
sleep(1)
pg.click(imp_coord_dict["first share"])
sleep(1)

for share in EQ_shares:






