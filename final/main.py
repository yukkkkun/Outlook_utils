# -*- coding:utf-8 -*-

import win32com.client
import tkinter as tk
import win32com.client
from datetime import datetime
import threading

#my module
import outlook_basic
import mode
import chrome



mode_selected = "ZAITAKU"

#######################################
###Utils
#######################################





#######################################
###Outlook基本コマンド
#######################################
# outlook = win32com.client.Dispatch("Outlook.Application")

# def sendEmail(dateandtime):
#     Msg = outlook.CreateItem(0) # Email
#     Msg.To = "test@test.com" # you can add multiple emails with the ; as delimiter. E.g. test@test.com; test2@test.com;
#     # Msg.CC = "test@test.com"

#     Msg.Subject = "名前_" + ""
#     Msg.Body = '''
    
#     在宅勤務をさせて頂きます。

    
#     '''
#     Msg.display(True)

# #   Msg.Send()

# def sendMeeting(dateandtime, body): 
#     str_dateandtime = str(dateandtime)   
#     appt = outlook.CreateItem(1) # AppointmentItem
#     appt.AllDayEvent = True
#     appt.Start = str_dateandtime # yyyy-MM-dd hh:mm
#     appt.Subject = "Subject of the meeting"
#     appt.Duration = 60 # In minutes (60 Minutes)
#     appt.Location = "Location Name"
#     appt.MeetingStatus = 1 # 1 - olMeeting; Changing the appointment to meeting. Only after changing the meeting status recipients can be added
    
#     appt.Recipients.Add("test@test.com") # Don't end ; as delimiter
    
#     appt.Body = body

#     appt.display(True)
#     #   appt.Save()
#     #   appt.Send()

# def add_outlook_schedule(start_time, duration, body):
# 	APPOINTMENT_ITEM = 1

# 	outlook = win32com.client.Dispatch("Outlook.Application")
# 	mapi = outlook.GetNamespace("MAPI")
# 	item = outlook.CreateItem(APPOINTMENT_ITEM)

# 	item.Start = start_time
# 	item.Duration = duration
# 	item.Subject = '在宅：電話番号'
# 	item.Body = body
# 	item.ReminderMinutesBeforeStart = 0
# 	item.ReminderSet = True
# 	item.Save()


# def sendRecurringMeeting():    
#     appt = outlook.CreateItem(1) # AppointmentItem
#     appt.Start = "2018-10-28 10:10" # yyyy-MM-dd hh:mm
#     appt.Subject = "Subject of the meeting"
#     appt.Duration = 60 # In minutes (60 Minutes)
#     appt.Location = "Location Name"
#     appt.MeetingStatus = 1 # 1 - olMeeting; Changing the appointment to meeting. Only after changing the meeting status recipients can be added

#     appt.Recipients.Add("test@test.com") # Don't end ; as delimiter

#     # Set Pattern, to recur every day, for the next 5 days
#     pattern = appt.GetRecurrencePattern()
#     pattern.RecurrenceType = 0
#     pattern.Occurrences = "5"

#     #   appt.Save()
#     #   appt.Send()


#######################################
###時短関数（こっからが本番）
#######################################

# def zaitaku(dateandtime):
    

#     subject = '在住'
#     body = '''
    
#     在宅勤務をさせていただきます。
    
    
#     '''

#     outlook_basic.sendMeeting(dateandtime, subject, body)

#     start_time = dateandtime + ' 8:30'
#     duration = 540
#     outlook_basic.add_outlook_schedule(start_time, duration, body)


#######################################
### カレンダーを作成するフレームクラス
#######################################
class mycalendar(tk.Frame):
    def __init__(self,master=None,cnf={},**kw):
        "初期化メソッド"
        import datetime
        tk.Frame.__init__(self,master,cnf,**kw)

        # 現在の日付を取得
        now = datetime.datetime.now()
        # 現在の年と月を属性に追加
        self.year = now.year
        self.month = now.month
        # 追記 https://teratail.com/questions/234639#reply-355304
        global YEAR, MONTH
        YEAR = str(self.year)
        MONTH = str(self.month)

        # frame_top部分の作成
        frame_top = tk.Frame(self)
        frame_top.pack(pady=5)
        self.previous_month = tk.Label(frame_top, text = "<", font = ("",14))
        self.previous_month.bind("<1>",self.change_month)
        self.previous_month.pack(side = "left", padx = 10)
        self.current_year = tk.Label(frame_top, text = self.year, font = ("",18))
        self.current_year.pack(side = "left")
        self.current_month = tk.Label(frame_top, text = self.month, font = ("",18))
        self.current_month.pack(side = "left")
        self.next_month = tk.Label(frame_top, text = ">", font = ("",14))
        self.next_month.bind("<1>",self.change_month)
        self.next_month.pack(side = "left", padx = 10)

        # frame_week部分の作成
        frame_week = tk.Frame(self)
        frame_week.pack()
        button_mon = d_button(frame_week, text = "Mon")
        button_mon.grid(column=0,row=0)
        button_tue = d_button(frame_week, text = "Tue")
        button_tue.grid(column=1,row=0)
        button_wed = d_button(frame_week, text = "Wed")
        button_wed.grid(column=2,row=0)
        button_thu = d_button(frame_week, text = "Thu")
        button_thu.grid(column=3,row=0)
        button_fri = d_button(frame_week, text = "Fri")
        button_fri.grid(column=4,row=0)
        button_sta = d_button(frame_week, text = "Sat", fg = "blue")
        button_sta.grid(column=5,row=0)
        button_san = d_button(frame_week, text = "Sun", fg = "red")#'San'→'Sun'と修正した
        button_san.grid(column=6,row=0)

        # frame_calendar部分の作成
        self.frame_calendar = tk.Frame(self)
        self.frame_calendar.pack()

        # 日付部分を作成するメソッドの呼び出し
        self.create_calendar(self.year,self.month)

    def create_calendar(self,year,month):
        "指定した年(year),月(month)のカレンダーウィジェットを作成する"

        # ボタンがある場合には削除する（初期化）
        try:
            for key,item in self.day.items():
                item.destroy()
        except:
            pass

        # calendarモジュールのインスタンスを作成
        import calendar
        cal = calendar.Calendar()
        # 指定した年月のカレンダーをリストで返す
        days = cal.monthdayscalendar(year,month)

        # 日付ボタンを格納する変数をdict型で作成
        self.day = {}
        # for文を用いて、日付ボタンを生成
        for i in range(0,42):
            c = i - (7 * int(i/7))
            r = int(i/7)
            try:
                # 日付が0でなかったら、ボタン作成
                if days[r][c] != 0:
                    self.day[i] = d_button(self.frame_calendar,text = days[r][c])
                    self.day[i].grid(column=c,row=r)
            except:
                """
                月によっては、i=41まで日付がないため、日付がないiのエラー回避が必要
                """
                break

    def change_month(self,event):
        # 押されたラベルを判定し、月の計算
        if event.widget["text"] == "<":
            self.month -= 1
        else:
            self.month += 1
        # 月が0、13になったときの処理
        if self.month == 0:
            self.year -= 1
            self.month = 12
        elif self.month == 13:
            self.year +=1
            self.month =1
        # frame_topにある年と月のラベルを変更する
        self.current_year["text"] = self.year
        self.current_month["text"] = self.month

        # 追記 https://teratail.com/questions/234639#reply-355304
        global YEAR, MONTH
        YEAR = str(self.year)
        MONTH = str(self.month)

        # 日付部分を作成するメソッドの呼び出し
        self.create_calendar(self.year,self.month)

# デフォルトのボタンクラス
class d_button(tk.Button):
    def __init__(self,master=None,cnf={},**kw):
        tk.Button.__init__(self,master,cnf,**kw)
        self.configure(font=("",14),height=2, width=4, relief="flat")
        self.bind('<Button-1>', callback)# 追記 https://teratail.com/questions/234639

# カレンダーの年月日を取得するコールバック関数
# 追記 https://teratail.com/questions/234639#reply-355304
def callback(event):
    selected_date = ''
    if event.widget['text'] not in ['Mon','Tue','Wed','Thu','Fri','Sat','Sun']:
        selected_date += YEAR + '-'
        selected_date += convert_in2_2bytes(MONTH) + '-'
        selected_date += convert_in2_2bytes(str(event.widget['text']))
        print(selected_date)



        if mode_selected == "ZAITAKU":
            mode.zaitaku(selected_date)


# 1桁の数字を2バイトに変換する関数
# 追記 https://teratail.com/questions/234639#reply-355304
def convert_in2_2bytes(str_number):
    if len(str_number) == 1:
        return '0' + str_number
    else:
        return str_number



if __name__=="__main__":
    
    def thread_chrome():
        chrome.open_with_extensions()

    # t1 = threading.Thread(target=thread_outlook)
    thread1 = threading.Thread(target=thread_chrome)

    # t1.start()
    thread1.start()

    # ルートフレームの定義   
    root = tk.Tk()
    root.title("Calendar App")
    mycal = mycalendar(root)
    mycal.pack()
    root.mainloop()    # sendMeeting()