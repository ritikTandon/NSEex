import ctypes
import math
import random
import time

from DataStructures import *


# def getIntersectionNode(headA, headB):
#     temp = headA
#     node_list = []
#     flag = False
#
#     while temp:
#         node_list.append(id(temp))
#         temp = temp.next
#
#     temp = headB
#
#     while temp:
#         if id(temp) in node_list:
#             return ctypes.cast(id(temp), ctypes.py_object).value
#
#         temp = temp.next
#
#     return None
#
#
# join = ListNode(99)
# nodeA = ListNode(1, ListNode(2, ListNode(3, join)))
# nodeB = ListNode(-1, ListNode(-2, ListNode(-3, join)))
#
# print(f"Intersected at: {getIntersectionNode(nodeA, nodeB).val}")

sh = """AARTIIND
ABB
ABCAPITAL
ABFRL
ADANIENT
ADANIPORTS
ALKEM
AMARAJABAT
AMBUJACEM
APLLTD
APOLLOHOSP
APOLLOTYRE
ASHOKLEY
ASTRAL
ATUL
AUBANK
AUROPHARMA
BAJAJFINSV
BAJFINANCE
BALKRISIND
BALRAMCHIN
BANDHANBNK
BANKBARODA
BATAINDIA
BEL
BHARATFORG
BIOCON
BRATANNIA
BSOFT
CANBK
CANFINHOME
CHAMBLFERT
CHOLAFIN
CIPLA
COFORGE
CONCOR
COROMANDEL
CROMPTON
CUMMINSIND
DABUR
DALBHARAT
DEEPAKNTR
DELTACORP
DIVISLAB
DIXON
DLF
DRREDDY
ESCORTS
EXIDEIND
GLENMARK
GLS
GNFC
GODREJCP
GODREJPROP
GRANULES
GRASIM
GUJGASLTD
HAL
HAVELLS
HCLTECH
HDFCAMC
HDFCLIFE
HINDALCO
HINDCOPPER
ICICIGI
ICICIPRULI
IEX
IGL
INDHOTEL
INDIACEM
INDIAMART
INDIGO
INDUSINDBK
INDUSTOWER
INTELLECT
IPCALAB
JINDALSTEL
JKCEMENT
JSWSTEEL
JUBLFOOD
KOTAKBANK
LALPATHLAB
LAURUSLABS
LICHSGFIN
LTIM
LTTS
LUPIN
M&MFIN
MANAPPURAM
MARICO
MCDOWELL-N
MCX
METROPOLIS
MFSL
MGL
MPHASIS
MUTHOOTFIN
NAM-INDIA
NAUKRI
NAVINFLUOR
NMDC
NTPC
OBEROIRLTY
PEL
PERSISTENT
PETRONET
PIDILITIND
POLYCAB
POWERGRID
RAIN
RAMCOCEM
RBLBANK
RECLTD
SBICARD
SBILIFE
SIEMENS
SRF
STAR
SUNPHARMA
SYNGENE
TATACOMM
TECHM
TORNTPHARM
TORNTPOWER
TRENT
TVSMOTOR
UBL
ULTRACEMCO
UPL
VEDL
VOLTAS
ZEEL
ZYDUSLIFE"""

# L = sh.split("\n")
# print(L)
# j = 1
# for i in L:
#     print(i, end=" ")
#     if j % 3 == 0:
#         print()
#     # time.sleep(1)
#     j += 1
# print(len(L))

# for i in range(0, 11):
#     print(i**7)
#
# for i in range(0, 11):
#     print(i**109)







y = round_up(12.16)
print(y)

x = Decimal("3.456")
(x * 2).quantize(Decimal('.05'), rounding=ROUND_UP) / 2






