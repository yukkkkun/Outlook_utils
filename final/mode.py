import win32com.client
import tkinter as tk
import win32com.client
from datetime import datetime

#my module
import outlook_basic
import chrome

def zaitaku(dateandtime):
    
    with open('contents.txt') as f:
        for line in f.readlines():
            globals()["{0:s}".format(line.split()[0])] = line.split()[2]

    sendto = globals()['SENDTO']
    subject = globals()['SUBJECT']
    body = globals()['BODY']
    start_time = globals()['START_TIME']
    duration = int(globals()['DURATION'][1:-1])

    print(type(dateandtime))
    print(start_time)
    start_time = str(dateandtime) + ' ' + start_time[1:-1]
    print(start_time)
    # duration = 540
    outlook_basic.add_outlook_schedule(start_time, duration, body)

    outlook_basic.sendMeeting(sendto ,dateandtime, subject, body)


if __name__=="__main__":
    zaitaku(20210217)