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
# import random
import shutil

from date_variables import yr, mnth, date


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

# import openpyxl as xl
# from openpyxl.styles import Font, Alignment, PatternFill
#
# wb = xl.load_workbook(r'C:\Users\admin\PycharmProjects\daily data\test.xlsx')
#
# s = wb['Sheet1']
#
# s.cell(1,1).fill = PatternFill("solid", "FFFF00")
#
# wb.save(r'C:\Users\admin\PycharmProjects\daily data\test.xlsx')

# def wordPattern(pattern: str, s: str) -> bool:
#     l = s.split(" ")
#     i = 0
#     d = {}
#
#     if len(s) != len(l):
#         return False
#
#     for ele in pattern:
#         try:
#             if ele not in d and l[i] not in d.values():
#                 d[ele] = l[i]
#
#             elif d[ele] != l[i]:
#                 return False
#
#             i += 1
#         except KeyError:
#             return False
#
#     return True
#
#
# print(wordPattern("abba", "dog cat cat dog"))

# def deleteGreatestValue(grid: list[list[int]]) -> int:
#     tot = 0
#     while grid != [[]]:
#         m = []
#
#         for l in grid:
#             if not l:
#                 return tot
#
#             m.append(l.pop(l.index(max(l))))
#
#         tot += max(m)
#
#     return tot
#
#
# print(deleteGreatestValue([[1, 2, 4], [3, 3, 1]]))


# def lengthOfLongestSubstring(s: str) -> int:
#     mx = []
#     l = 0
#     vis = []
#     i = 0
#
#     if s == "":
#         return 0
#
#     if len(s) == 1:
#         return 1
#
#     if len(s) == 2:
#         return 2 if s[0] != s[1] else 1
#
#     while i < len(s):
#         char = s[i]
#         if char in vis:
#             l = len(vis)
#             mx.append(l)
#             vis = []
#             i = i - l+1
#             continue
#
#         else:
#             vis.append(char)
#
#         i += 1
#
#     if len(mx) > 0:
#         return max(max(mx), len(vis))
#     else:
#         return len(vis)
#
#
#
# print(lengthOfLongestSubstring("bwf"))


# def lengthOfLastWord(s: str) -> int:
#     i = len(s) - 1
#     enc = False
#     count = 0
#
#     if s == " ":
#         return 0
#
#     # if " " not in s:
#     #     return len(s)
#
#     while True:
#         if i == -1:
#             return count
#
#         if s[i] == " " and not enc:
#             i -= 1
#
#         elif s[i] != " ":
#             enc = True
#             count += 1
#             i -= 1
#
#         elif s[i] == " ":
#             return count

# def lengthOfLastWord(s: str) -> int:
#     l = len(s) - 1
#     count = 0
#
#     if s == " ":
#         return 0
#
#     for i in range(l, -1, -1):
#         if s[i] != " ":
#             count += 1
#         elif s[i] == " " and count > 0:
#             return count
#     return count
#
#
# print(lengthOfLastWord("a msas "))


# class ListNode:
#     def __init__(self, val=0, next=None):
#         self.val = val
#         self.next = next
#
#
# def addTwoNumbers(l1: list, l2: list): # list
#     p1 = 0
#     p2 = 0
#     i = 0
#     sum = 0
#     carry = 0
#
#     res = []
#
#     while True:
#         if p1 == len(l1):
#             while p2 < len(l2):
#                 sum = l2[p2] + carry
#                 if sum > 9:
#                     sum %= 10
#                     carry = 1
#                 else:
#                     carry = 0
#                 res.append(sum)
#                 p2 += 1
#
#             break
#
#         if p2 == len(l2):
#             while p1 < len(l1):
#                 sum = l1[p1]+carry
#                 if sum > 9:
#                     sum %= 10
#                     carry = 1
#                 else:
#                     carry = 0
#                 res.append(sum)
#                 p1 += 1
#
#             break
#
#         sum = l1[p1] + l2[p2] + carry
#
#         if sum > 9:
#             sum = sum % 10
#             carry = 1
#         else:
#             carry = 0
#
#         res.append(sum)
#
#         i += 1
#         p1 += 1
#         p2 += 1
#
#     if carry == 1:
#         res.append(carry)
#
#     return res


# print(addTwoNumbers([1,1], [1,1,1]))


# def addTwoNumbers(l1, l2):      # linked list
#     p1 = l1
#     p2 = l2
#
#     sum = 0
#     carry = 0
#
#     res = []
#
#
#     # print(len1, len2)
#
#     while True:
#         if p1 is None and p2 is None:
#             break
#
#         if p1 is None and p2 is not None:
#             while p2 is not None:
#                 sum = p2.val + carry
#                 if sum > 9:
#                     sum %= 10
#                     carry = 1
#                 else:
#                     carry = 0
#                 res.append(ListNode(sum))
#                 p2 = p2.next
#
#             break
#
#         if p2 is None and p1 is not None:
#             while p1 is not None:
#                 sum = p1.val + carry
#                 if sum > 9:
#                     sum %= 10
#                     carry = 1
#                 else:
#                     carry = 0
#                 res.append(ListNode(sum))
#                 p1 = p1.next
#
#             break
#
#         sum = p1.val + p2.val + carry
#
#         if sum > 9:
#             sum = sum % 10
#             carry = 1
#         else:
#             carry = 0
#
#         res.append(ListNode(sum))
#
#         p1 = p1.next
#         p2 = p2.next
#
#     if carry == 1:
#         res.append(ListNode(1))
#
#     for i in range(len(res)-1):
#         if res[i].next is None:
#             res[i].next = res[i+1]
#
#     return res[0]
#
#
# l2 = ListNode(2, next=ListNode(4, next=None))
# l1 = ListNode(5, next=ListNode(6, next=ListNode(4, next=None)))
#
# print(addTwoNumbers(l1, l2))


# l2 = ListNode(2, next=ListNode(4, next=None))
# l1 = ListNode(5, next=ListNode(6, next=ListNode(4, next=None)))
#
# print(addTwoNumbers(l1, l2))

# def isPalindrome(s: str) -> bool:
#     new = ""
#     valid = ['a', 'A', 'b', 'B', 'c', 'C', 'd', 'D', 'e', 'E', 'f', 'F', 'g', 'G', 'h', 'H', 'i', 'I', 'j', 'J', 'k',
#              'K', 'l', 'L', 'm', 'M', 'n', 'N', 'o', 'O', 'p', 'P', 'q', 'Q', 'r', 'R', 's', 'S', 't', 'T', 'u', 'U',
#              'v', 'V', 'w', 'W', 'x', 'X', 'y', 'Y', 'z', 'Z']
#
#     if s == " ":
#         return True
#
#     if len(s) < 3:
#         return True
#
#     b = False
#     for ch in s:
#         if ch in valid:
#             new += ch.lower()
#
#     if len(new) == 1:
#         return True
#
#     i = 0
#     j = len(new) - 1
#
#     while i < j:
#         first = new[i]
#         last = new[j]
#
#         if first == last:
#             b = True
#
#         else:
#             b = False
#             return b
#
#         i += 1
#         j -= 1
#
#     return b
#
#
# print(isPalindrome(".,"))


# def convertToTitle(columnNumber: int) -> str:
#     res = ''
#     alpha = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
#     while columnNumber > 26:
#         rem = columnNumber % 26
#         columnNumber = columnNumber // 26
#         if rem == 0:
#             columnNumber -= 1
#         res = alpha[rem - 1] + res
#
#     res = alpha[columnNumber - 1] + res
#
#     return res
#
#
# print(convertToTitle(11))
#
#
# def titleToNumber(columnTitle: str) -> int:
#     i = len(columnTitle) - 1
#     tot = 0
#     alpha = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
#     for s in columnTitle:
#         tot += (26 ** i) * (alpha.index(s)+1)
#
#         i -= 1
#
#     return tot
#
# print(titleToNumber("K"))

# def numIdenticalPairs(nums):
#     hashMap = {}
#     res = 0
#     for number in nums:
#         if number in hashMap:
#             res += hashMap[number]
#             hashMap[number] += 1
#         else:
#             hashMap[number] = 1
#     return res
#
#
# l = []
#
# for i in range(1000000):
#     l.append(random.randint(0, 10000))
#
# # print(l)
# print(numIdenticalPairs(l))

# def reverseString(s) -> None:
#     """
#     Do not return anything, modify s in-place instead.
#     """
#     for i in range(len(s)-1):
#         t = s.pop(0)
#         s.insert(len(s)-i, t)
#
#         # print(i)


# def shuffle(nums: list, n: int) -> list[int]:
#     for i in range(n):
#         nums[i] += nums[n + i] * 10000
#
#     i = n - 1
#
#     # after this iteration, the list will be [40001, 50002, 60003, 4, 5, 6] essentially storing both numbers in 1 number and preserving it
#
#     while i >= 0:       # here, we are going in reverse and taking the mod/integer division to extract the numbers
#         nums[2 * i + 1] = nums[i] // 10000  # for even digits, the y1 part
#         nums[2 * i] = nums[i] % 10000       # for odd digits, the x1 part
#         i -= 1
#
#     return nums
#
#
# print(shuffle([1, 2, 3, 4, 5, 6], 3))
#
# class ListNode:
#     def __init__(self, x, next=None):
#         self.val = x
#         self.next = next
#
#
# def deleteNode(node):
#     while node and node.next:
#         node.val = node.next.val
#         node.next = node.next.next
#         node = node.next
#
#
# l = ListNode(4, ListNode(5, ListNode(1, ListNode(9, None))))
#
# deleteNode(ListNode(1, ListNode(9, None)))
# print(l)


# def rotate(matrix):
#     # reverse
#     l = 0
#     r = len(matrix) - 1
#     while l < r:
#         matrix[l], matrix[r] = matrix[r], matrix[l]
#         l += 1
#         r -= 1
#     # transpose
#     for i in range(len(matrix)):
#         for j in range(i):
#             matrix[i][j], matrix[j][i] = matrix[j][i], matrix[i][j]
# 
#
# # rotate([[1,2,3],[4,5,6],[7,8,9]])
# s = "APOLLO BAJFINSV BAJFIN BARODA BN COALIND DLF EICHER FEDBANK HCL HDFC ICICI INDUSIND INFY JIND M&M M&MFIN NIFTY REL SBIN SUNTV TCON TM TP TS TITAN ULTRA VEDL"
#
# print(s.split(" "))

# shutil.copy(rf"E:\Daily Data work\hourlys 30 minute FO\{yr}\{mnth}\{date}\NIFTY.xls", rf"E:\Daily Data work\hourlys 30 minute CASH\{yr}\{mnth}\{date}")
# shutil.copy(rf"E:\Daily Data work\hourlys 30 minute FO\{yr}\{mnth}\{date}\BN.xls", rf"E:\Daily Data work\hourlys 30 minute CASH\{yr}\{mnth}\{date}")

# def hammingWeight(n: int) -> int:
#     count = 0
#
#     while n:
#         count += n%2
#         n = n >> 1
#
#     return count
# print(hammingWeight(0000000000000000000000010000000))


# def moveZeroes(nums) -> None:
#     slow = 0
#     for fast in range(len(nums)):
#         if nums[fast] != 0 and nums[slow] == 0:
#             nums[slow], nums[fast] = nums[fast], nums[slow]
#
#         # wait while we find a non-zero element to
#         # swap with you
#         if nums[slow] != 0:
#             slow += 1
#
#
# l = [0, 1, 0, 3, 12]
# moveZeroes(l)
# print(l)


# def buildArray(nums):   # O(1) - In-place
#     for i in range(len(nums)):
#         nums[i] = (nums[nums[i]] % 10000) * 10000 + nums[i]
#
#     return [num // 10000 for num in nums]
#
#
#
# print(buildArray([0, 2, 1, 5, 3, 4]))

# Code to Measure time taken by program to execute.
import time
begin = time.time()


def bestClosingTime(customers: str) -> int:
    h = m = s = 0
    for i, ch in enumerate(customers):  # [1] compute running profit where
        s += (ch == "Y") * 2 - 1  # we add +1 for Y, -1 for N
        if s > m:  # [2] keep track of the maximal
            m, h = s, i + 1  # profit and its hour

    return h


print(bestClosingTime("YYNY"))
end = time.time()
print(f"Total runtime of the program is {end - begin}s")