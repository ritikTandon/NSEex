import math
import random
from DataStructures import *


def getIntersectionNode(headA, headB):
    temp = headA
    node_list = []

    while temp:
        node_list.append(temp)
        temp = temp.next

    temp = headB

    while temp:
        if temp in node_list:
            return node_list.index(temp)

        temp = temp.next

    return None


getIntersectionNode()








