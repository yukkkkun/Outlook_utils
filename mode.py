import win32com.client
import tkinter as tk
import win32com.client
from datetime import datetime

#my module
import outlook_basic


def zaitaku(dateandtime):
    

    subject = '在住'
    body = '''
    
    在宅勤務をさせていただきます。
    
    
    '''

    outlook_basic.sendMeeting(dateandtime, subject, body)

    start_time = dateandtime + ' 8:30'
    duration = 540
    outlook_basic.add_outlook_schedule(start_time, duration, body)
