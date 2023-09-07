import random


class TreeNode:
    def __init__(self, val=0, left=None, right=None):
        self.val = val
        self.left = left
        self.right = right


class ListNode:
    def __init__(self, val=0, next=None):
        self.val = val
        self.next = next

    def __str__(self):
        l = []
        temp = self

        while temp:
            l.append(temp.val)
            temp = temp.next

        # print(l)
        return str(l)

    def print(self):
        l = []
        temp = self

        while temp:
            l.append(temp.val)
            temp = temp.next

        print(l)


def create_linked_list(val_list: list=None, count: int=0):
    if val_list is None:
        val_list = []
        for i in range(count):
            val_list.append(random.randint(0, 10))

    linked_list = ListNode()
    temp = linked_list

    for i in range(len(val_list)):
        temp.val = val_list[i]
        temp.next = ListNode() if i < len(val_list)-1 else None
        temp = temp.next

    return linked_list
