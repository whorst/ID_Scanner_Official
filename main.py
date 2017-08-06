from tkinter import *
import xlrd
import os
import time


class GetBrother:
    def __init__(self):
        self.member_dict = {}
        self.people_present= {}
        self.brother_excel_file = xlrd.open_workbook("Brother Contact Information Final.xlsx").sheet_by_index(0)
        self.brother_count = len(self.brother_excel_file.col_values(0))

    def populate_member_dict(self):
        i = 1
        while (i < self.brother_count):
            list_brother_name_and_number = self.brother_excel_file.row_values(i)
            brother_name = list_brother_name_and_number[0]
            brother_id_number = str(int(list_brother_name_and_number[4]))
            self.member_dict[brother_id_number] = brother_name
            i += 1

GetBrother = GetBrother()
GetBrother.populate_member_dict()
#print(GetBrother.member_dict)



print(time.strftime("%I:%M:%S %p"))

# while(True):
#     card_code = input("SHOW ME WHAT YOU GOT I WANT TO SEE WHAT YOU GOT\n")
#     for k,v in GetBrother.member_dict.items():
#         if k in card_code[26:]:
#             print(v)



    # global_id_number = card_code[29:]
    # global_id_number = global_id_number.replace("?","")



    # if member_dict[global_id_number] != None:
    #     print(member_dict[global_id_number]+"\n")

