# -*- coding:utf-8 -*-

import win32com.client
import tkinter as tk
import win32com.client
from datetime import datetime



#######################################
###Utils
#######################################





#######################################
###Outlook基本コマンド
#######################################
outlook = win32com.client.Dispatch("Outlook.Application")

def sendEmail(sendto, subject, body):
    Msg = outlook.CreateItem(0) # Email
    Msg.To = sendto # you can add multiple emails with the ; as delimiter. E.g. test@test.com; test2@test.com;
    # Msg.CC = "test@test.com"

    Msg.Subject = subject
    Msg.Body = body
    Msg.display(True)

#   Msg.Send()

def sendMeeting(sendto, dateandtime, subject, body): 
    str_dateandtime = str(dateandtime)   
    appt = outlook.CreateItem(1) # AppointmentItem
    appt.AllDayEvent = True
    appt.Start = str_dateandtime # yyyy-MM-dd hh:mm
    appt.Subject = subject
    appt.Duration = 60 # In minutes (60 Minutes)
    appt.Location = "Location Name"
    appt.MeetingStatus = 1 # 1 - olMeeting; Changing the appointment to meeting. Only after changing the meeting status recipients can be added
    
    appt.Recipients.Add(sendto) # Don't end ; as delimiter
    
    appt.Body = body

    appt.display(True)
    #   appt.Save()
    #   appt.Send()

def add_outlook_schedule(start_time, duration, body):
	APPOINTMENT_ITEM = 1

	outlook = win32com.client.Dispatch("Outlook.Application")
	mapi = outlook.GetNamespace("MAPI")
	item = outlook.CreateItem(APPOINTMENT_ITEM)

	item.Start = start_time
	item.Duration = duration
	item.Subject = '在宅：電話番号'
	item.Body = body
	item.ReminderMinutesBeforeStart = 0
	item.ReminderSet = True
	item.Save()


def sendRecurringMeeting():    
    appt = outlook.CreateItem(1) # AppointmentItem
    appt.Start = "2018-10-28 10:10" # yyyy-MM-dd hh:mm
    appt.Subject = "Subject of the meeting"
    appt.Duration = 60 # In minutes (60 Minutes)
    appt.Location = "Location Name"
    appt.MeetingStatus = 1 # 1 - olMeeting; Changing the appointment to meeting. Only after changing the meeting status recipients can be added

    appt.Recipients.Add("test@test.com") # Don't end ; as delimiter

    # Set Pattern, to recur every day, for the next 5 days
    pattern = appt.GetRecurrencePattern()
    pattern.RecurrenceType = 0
    pattern.Occurrences = "5"

    #   appt.Save()
    #   appt.Send()




if __name__=="__main__":

    # # ルートフレームの定義      
    # root = tk.Tk()
    # root.title("Calendar App")
    # mycal = mycalendar(root)
    # mycal.pack()
    # root.mainloop()

    # sendMeeting()
    pass