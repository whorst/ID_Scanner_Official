import tkinter
from datetime import datetime
import xlrd
import xlsxwriter
import time
import re

'''
Still need to
    -Correct No Statements with something tangible
    -Get the date squared away
    -Write intelligible comments
    -learn how to bundle the app
    -Fuck Connor
'''

#00000000000000000000000000000617295?

class MyFirstGUI:
    def __init__(self, master):
        self.master = master
        self.xnum = 1
        self.ynum = 6
        self.master = master
        self.labels_to_delete = []

        master.title("The ΣΑΕ Attendance Software")

        self.entry_excel_title = tkinter.Entry(master)
        self.entry_excel_title.grid(row=1, column=1, pady=(30,7))

        self.label_excel_title = tkinter.Label(master, text="Enter Document Name")
        self.label_excel_title.grid(row=2, column=1)

        self.entry_start_time = tkinter.Entry(master)
        self.entry_start_time.grid(row=1, column=3, pady=(30,7))

        self.label_start_time = tkinter.Label(master, text="Enter Event Start Time")
        self.label_start_time.grid(row=2, column=3)

        self.button_submit_documnet = tkinter.Button(master, text="Submit Document")
        self.button_submit_documnet.grid(row=2, column=2)
        self.button_submit_documnet.bind("<Button>",lambda event, : self.export_data())

        tkinter.Label(master, text="").grid(row=3, column=1, padx=75, pady=70)
        tkinter.Label(master, text="Please Swipe Your ID").grid(row=3, column=2, padx=75, pady=70)
        tkinter.Label(master, text="").grid(row=3, column=3, padx=70, pady=70)

        self.entry_card_number = tkinter.Entry(master)
        self.entry_card_number.grid(row=4, column=2, pady=10)

        self.button_submit_id_number = tkinter.Button(master, text="Submit", width=16)
        self.button_submit_id_number.grid(row=5, column=2)
        self.button_submit_id_number.bind("<Button>",lambda event, : self.add_to_grid())

    def get_document_title(self):
        return self.entry_excel_title.get()

    def get_start_time(self):
        arrival_time = self.entry_start_time.get()
        time_array = arrival_time.split(":")
        if "pm" in time_array[1]:
            time_array[0] = str(int(time_array[0])+12)
        time_array[1] = re.sub("[^0-9]", "",time_array[1])
        self.event_time_start = time_array[0] + ":" + time_array[1]
        #print(self.event_time_start)
        return self.event_time_start

    def verify_document_title(self):
        '''
        This simply makes sure that the user has entered the excel title correctly:
        '''
        not_allowed_list=[]
        excel_title = self.entry_excel_title.get()
        if (("~" in excel_title) or ("#" in excel_title) or ("%" in excel_title) or ("&" in excel_title) or ("*" in excel_title) or ("{" in excel_title) or ("}" in excel_title) or ("\\" in excel_title) or (":" in excel_title) or ("<" in excel_title)  or (">" in excel_title)):
            print("No")
            return False
        else:
            return True

    def verify_card_input(self):
        card_number = self.entry_card_number.get()
        if len(card_number) == 36:
            return True
        else:
            return False

    def verify_time(self):
        '''
        This simply makes sure that the user has entered the time they want to begin at correctly:
        '''
        arrival_time = self.entry_start_time.get()
        print()
        print((len(arrival_time)< 6))
        #print(("am" not in arrival_time.lower()) and ("pm" not in arrival_time.lower()))
        print()
        if((len(arrival_time)< 6) and (("am" not in arrival_time.lower()) or ("pm" not in arrival_time.lower()))):
            return False
        else:
            return True

    def arrived_on_time(self, time_to_compare):
        '''
        This method will be confusing. Split the 24 hour times up into arrays, with the first element being the hour
        and the second element being the minute. Compare the hours. Then the minutes.
        '''
        print("Event started at"+self.event_time_start)
        print("Person arrived at"+time_to_compare)
        event_start_time_array = self.event_time_start.split(":")
        time_to_compare_array = time_to_compare.split(":")

        print(time_to_compare_array, event_start_time_array)

        if int(time_to_compare_array[0]) < int(event_start_time_array[0]):
            #print(1)
            return True

        if int(time_to_compare_array[0]) > int(event_start_time_array[0]):
            #print(2)
            return False

        if int(time_to_compare_array[0]) == int(event_start_time_array[0]):
            if(int(time_to_compare_array[1]) < int(event_start_time_array[1])):
                #print(3)
                return True
            if(int(time_to_compare_array[1]) > int(event_start_time_array[1])):
                #print(4)
                return False
            if(int(time_to_compare_array[1]) == int(event_start_time_array[1])):
                #print(5)
                return True

    def export_data(self):

        self.present_column_num = 0
        self.present_row_num = 0

        self.absent_column_num = 0
        self.absent_row_num = 0

        #print(self.verify_document_title(), self.verify_time())
        GetBrother.get_not_present()

        if (self.verify_document_title() and self.verify_time()):

            print("Yes")

            event_start = self.get_start_time()
            event_title = self.get_document_title()+".xlsx"

            workbook = xlsxwriter.Workbook(event_title)

            #Create the different worksheets
            present_worksheet = workbook.add_worksheet("Present People")
            absent_worksheet = workbook.add_worksheet("Absent People")

            #Set the width of the columns
            present_worksheet.set_column("A:A", 30)
            present_worksheet.set_column("B:B", 30)
            absent_worksheet.set_column("A:A", 30)
            absent_worksheet.set_column("B:B", 30)

            #Write the brother's name and the time they've arrived to the relevant columns and rows
            present_worksheet.write_string(self.present_row_num, self.present_column_num, "Brother Name")
            present_worksheet.write_string(self.present_row_num, self.present_column_num+1, "Time Arrived")
            absent_worksheet.write_string(self.absent_row_num, self.absent_column_num, "Brother Name")
            absent_worksheet.write_string(self.absent_row_num, self.absent_column_num + 1, "Time Arrived")

            self.present_row_num+=1
            self.absent_row_num+=1
            #present_column_num+=1

            for key_brother_name, value_time in GetBrother.people_present.items():
                twelve_hour_time = datetime.strptime(value_time, "%H:%M")
                value_time = twelve_hour_time.strftime("%I:%M %p")
                #WOOOOOOOWWWWW THIS IS FUUUUCKED
                #print(self.arrived_on_time(value_time))
                if(self.arrived_on_time(value_time)):
                    present_worksheet.write_string(self.present_row_num, self.present_column_num, key_brother_name)
                    present_worksheet.write_string(self.present_row_num, self.present_column_num+1, value_time)
                    self.present_row_num+=1

                else:
                    absent_worksheet.write_string(self.absent_row_num, self.absent_column_num, key_brother_name)
                    absent_worksheet.write_string(self.absent_row_num, self.absent_column_num + 1, value_time)
                    self.absent_row_num+=1

            for key_absent_brother, value_absent_time in GetBrother.people_not_present.items():
                absent_worksheet.write_string(self.absent_row_num, self.absent_column_num, key_absent_brother)
                absent_worksheet.write_string(self.absent_row_num, self.absent_column_num + 1, value_absent_time)
                self.absent_row_num += 1

            workbook.close()

            self.entry_start_time.delete(0, 'end')
            self.entry_excel_title.delete(0, 'end')

            for label in self.labels_to_delete:
                label.destroy()
        else:
            print("NOOOOOO")

    def add_to_grid(self):

        if(self.verify_card_input()):

            card_number = self.entry_card_number.get()

            for id_number_key, brother_name_value in GetBrother.member_dict.items():
                #Check if the card you've swiped is in the excel document
                if id_number_key in card_number[26:]:
                    #If the brother isn't on the grid put him there. Otherwise don't
                    if ((brother_name_value in GetBrother.people_present) == False):
                        GetBrother.people_present[brother_name_value] = time.strftime("%H:%M")
                        brother_name_label = tkinter.Label(self.master, text=brother_name_value)
                        self.labels_to_delete.append(brother_name_label)
                        self.entry_card_number.delete(0, 'end')
                        brother_name_label.grid(row=self.ynum, column=self.xnum)
                        self.xnum += 1
                        if (self.xnum == 4):
                            self.xnum = 1
                            self.ynum = self.ynum + 1
                    elif((brother_name_value in GetBrother.people_present) == True):
                        self.entry_card_number.delete(0, 'end')

        else:
            print("No")

            # print(self.xnum, self.ynum)

class GetBrother:
    def __init__(self):
        self.member_dict = {}
        self.people_present= {}
        self.people_not_present = {}
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

    def get_not_present(self):
        for key_number, value_brother_name in self.member_dict.items():
            if(value_brother_name not in self.people_present):
                self.people_not_present[value_brother_name] = "Did Not Show"


GetBrother = GetBrother()
GetBrother.populate_member_dict()
#print(GetBrother.member_dict)

root = tkinter.Tk()
root.minsize(width=600, height=600)
my_gui = MyFirstGUI(root)
root.mainloop()