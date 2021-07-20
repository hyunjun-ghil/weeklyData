# Weekl Data Program
# 기능 : 날짜 선택 기능 (From ~ To)
#       쿼리 결과 받아와서 엑셀 데이터로 출력
#       엑셀 데이터 출력 시 경로 지정
#       Progress bar


import tkinter.messagebox as msgbox
from tkinter.filedialog import *
from tkinter import font
from tkinter import filedialog

import datetime
from datetime import timedelta
from tkcalendar import Calendar
import pymssql
import pandas as pd

try:
    import tkinter as tk
    from tkinter import ttk
except ImportError:
    import Tkinter as tk
    import ttk


dir_path = ""
curdatetime = datetime.date.today()
curdatetime2 = datetime.date.today()
todaydate = str(curdatetime)
todaydate2 = str(curdatetime)
filename = ""

#################### 함수 정의부 시작 ##########################

# 달력 보여주는 함수
def show_calendar():
    def print_sel():
        global curdatetime
        print(cal.selection_get())
        cal.see(datetime.date(year=today.year, month=today.month, day=today.day))
        label1.config(text=cal.selection_get())
        curdatetime = cal.selection_get()
        top.destroy()

    top = tk.Toplevel(title_frame)

    import datetime
    today = datetime.date.today()
    print(today)

    cal = Calendar(top, font="Arial 14", selectmode='day', locale='en_US',
                   disabledforeground='red', cursor="hand1", year=today.year, month=today.month, day=today.day)
    cal.pack(fill="both", expand=True)
    ttk.Button(top, text="ok", command=print_sel).pack()

# 하루 전 날짜로 이동하는 함수
def dateleft():
    global curdatetime
    global todaydate
    yesterday = curdatetime - timedelta(days=1)
    curdatetime = yesterday
    todaydate = str(yesterday)
    label1.config(text=todaydate)


# 하루 후 날짜로 이동하는 함수
def dateright():
    global curdatetime
    global todaydate

    yesterday = curdatetime + timedelta(days=1)
    curdatetime = yesterday
    todaydate = str(yesterday)
    label1.config(text=todaydate)


# 달력 보여주는 함수2
def show_calendar2():
    def print_sel():
        global curdatetime2
        print(cal.selection_get())
        cal.see(datetime.date(year=today.year, month=today.month, day=today.day))
        label2.config(text=cal.selection_get())
        curdatetime2 = cal.selection_get()
        top.destroy()

    top = tk.Toplevel(title_frame)

    import datetime
    today = datetime.date.today()
    print(today)

    cal = Calendar(top, font="Arial 14", selectmode='day', locale='en_US',
                   disabledforeground='red', cursor="hand1", year=today.year, month=today.month, day=today.day)
    cal.pack(fill="both", expand=True)
    ttk.Button(top, text="ok", command=print_sel).pack()

# 하루 전 날짜로 이동하는 함수
def dateleft2():
    global curdatetime2
    global todaydate2
    yesterday = curdatetime2 - timedelta(days=1)
    curdatetime2 = yesterday
    todaydate2 = str(yesterday)
    label2.config(text=todaydate2)


# 하루 후 날짜로 이동하는 함수
def dateright2():
    global curdatetime2
    global todaydate2

    yesterday2 = curdatetime2 + timedelta(days=1)
    curdatetime2 = yesterday2
    todaydate2 = str(yesterday2)
    label2.config(text=todaydate2)

# 엑셀 출력 함수
def exportData():
    try:
        if dir_path == '' or dir_path is None:
            msgbox.showinfo("알림", "저장 경로를 설정해주세요.")
            return

        global curdatetime
        global curdatetime2

        print("btn_1 Clicked")

        wFromDate = str(curdatetime)
        wToDate = str(curdatetime2)

        print(wFromDate)
        print(wToDate)

        FromDate = wFromDate
        ToDate = wToDate

        try:
            conn = pymssql.connect(host=r"", user='', password='', database='',
                                   charset='utf8')
            # Connection 으로부터 Cursor 생성
            cursor = conn.cursor()
            # SQL문 실행

            query = """ """

            query2 = """ """

            p_var2.set(25)  # progress bar 값 설정
            progressbar2.update()  # ui 업데이트

            df = pd.read_sql_query(query, conn)
            result_list.insert(END, "Raw데이터 생성 완료")
            p_var2.set(50)  # progress bar 값 설정
            progressbar2.update()  # ui 업데이트

            df2 = pd.read_sql_query(query2, conn)
            result_list.insert(END, "조치정보 생성")
            p_var2.set(100)  # progress bar 값 설정
            progressbar2.update()  # ui 업데이트

            writer = pd.ExcelWriter(
                dir_path + '/(대외비)' + FromDate[0:4] + '년' + FromDate[5:7] + FromDate[8:10] + '~' + ToDate[5:7] + ToDate[
                                                                                                        8:10] + '주간.xlsx',
                engine='xlsxwriter')

            df.to_excel(writer, sheet_name='Raw데이터', index=False, na_rep='NULL')
            print('Raw데이터 생성 완료')

            df2.to_excel(writer, sheet_name='조치정보', index=False, na_rep='NULL')
            print('조치정보 생성 완료')

            writer.save()

        except Exception as e:
            print('예외', e)

        # 연결 끊기
        conn.close()

        msgbox.showinfo("알림", "저장되었습니다.")

    except Exception as e:
        print('예외', e)


# 파일 경로 출력 팝업
def setPath():

    global dir_path

    dir_path = filedialog.askdirectory(parent=root, initialdir="/", title='저장 경로를 선택하세요.')
    path_txt.insert(END, dir_path)


#################### 함수 정의부 종료 ##########################


####################### Layout 시작 ###########################
root = Tk()
root.title("주간현황 추출 프로그램")

filename = str(curdatetime) + '.txt'

# 제목 프레임 ( 날짜이동 버튼 ) 날짜, 달력, 저장 열기 버튼
title_frame = LabelFrame(root, text="날짜 지정")
title_frame.pack()

titleFont = font.Font(family="맑은 고딕", size=10, weight="bold")

titleLabel = Label(title_frame, padx=5, pady=5, width=10, text="시작 날짜")
titleLabel.grid(row=0, column=0, columnspan=3)

titleLabel2 = Label(title_frame, padx=5, pady=5, width=10, text="종료 날짜")
titleLabel2.grid(row=0, column=4, columnspan=3)

date_left = Button(title_frame, text="◀", command=dateleft)
date_left.grid(row=1, column=0, sticky=E + W)

label1 = Button(title_frame, text=todaydate, width=10, command=show_calendar, font=titleFont)
label1.grid(row=1, column=1, sticky=E + W)

date_right = Button(title_frame, text="▶", command=dateright)
date_right.grid(row=1, column=2, sticky=E + W)

label3 = Label(title_frame, text='~', width=2, font=titleFont)
label3.grid(row=1, column=3, sticky=E + W)

date_left2 = Button(title_frame, width=2, text="◀", command=dateleft2)
date_left2.grid(row=1, column=4, sticky=E + W)

label2 = Button(title_frame, text=todaydate, width=10, command=show_calendar2, font=titleFont)
label2.grid(row=1, column=5, sticky=E + W)

date_right2 = Button(title_frame, width=2, text="▶", command=dateright2)
date_right2.grid(row=1, column=6, sticky=E + W)



context_frame = LabelFrame(root, text="출력")
context_frame.pack()

set_path_btn = Button(context_frame, padx=3, pady=3, width=40, text="저장 경로 지정", command=setPath, fg="black", bg="skyblue")
set_path_btn.grid(row=2, column=0, columnspan=7)

path_txt = Text(context_frame, width=40, height=1)
path_txt.grid(row=3, column=0, columnspan=7)

export_btn = Button(context_frame, padx=3, pady=3, width=40, text="엑셀 추출", command=exportData, fg="black", bg="skyblue")
export_btn.grid(row=4, column=0, columnspan=7)

result_list = Listbox(context_frame, width=40, height=2)
result_list.grid(row=5, column=0, columnspan=7, ipadx=2)

p_var2 = DoubleVar()
progressbar2 = ttk.Progressbar(context_frame, maximum=100, length=150, variable=p_var2)
progressbar2.grid(row=6, column=0, columnspan=7, ipadx=2, sticky=N + E + W + S)

######################################################################################

root.resizable(False, False)
root.mainloop()

