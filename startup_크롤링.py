# *********************************************
import tkinter as tk
import requests
import openpyxl
import datetime
import time
import sys

from tkinter import *
from tkinter import messagebox
from bs4 import BeautifulSoup
from urllib import request
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE

# *********************************************

# UI 설정
root = tk.Tk()
root.title("page 선택")
root.geometry("470x100+200+200")
root.resizable(False, False)

def start():
    global start_page, end_page
    start_page = int(input_start.get())
    end_page = int(input_end.get()) + 1

    if end_page < 1 :
        tk.messagebox.showinfo("경고","0보다 큰숫자를 입력하세요.")
    else:
        root.destroy()

lbl1 = Label(root, text = " 시작 번호 ", font = "NanumGothic 8")
# lbl1.grid(row=1, column=0)
lbl1.place(x=10, y=20)

input_start = Entry(root,width=7)
# input_page.grid(row=1, column=1)
input_start.place(x=300, y=20)

lbl2 = Label(root, text = " 종료 번호  ", font = "NanumGothic 8")
# lbl2.grid(row=2, column=0)
lbl2.place(x=10, y=50)

input_end = Entry(root,width=7)
# input_find.grid(row=2, column=1)
input_end.place(x=300, y=50)

btn = Button(root, text ="OK", command = start, width=3, height=1)
# btn.grid(row=2, column=2)
btn.place(x=420, y=50)

lbl3 = Label(root, text = " ※ 문의처 : hazelnut@kodit.co.kr ", font = "NanumGothic 8")
lbl3.place(x=10, y=80)

root.mainloop()

# *********************************************

# 엑셀저장 설정
xl = openpyxl.Workbook()
sheet = xl.active
sheet.title = "STARTUP WIKI"

# 출력변수 설정
list = []
sheet.append(["index","표제어","내용"])

for i in range(start_page,end_page) :

    # 홈페이지 접속
    
    adr = "http://startup-wiki.kr/archives/" + str(i)
    page = requests.get(adr)
    raw = page.content  
    html = BeautifulSoup(raw, 'html.parser')

    try:    
        title = html.find('header',class_='entry-header').text.strip()
        print(title)
        content = html.find('div',class_='entry-content').find('p', style = "text-align: justify;").text.strip()
        print(content)
        list.append([i,title,content])
        sheet.append([i,title,content])
    except AttributeError :
        print(str(i)+" Exception 2 : No Number Page")
        time.sleep(1) 


# 엑셀 저장
dt = datetime.datetime.now()
savetime = dt.strftime("_%y_%m_%d")
xl.save('START UP WIKI' + savetime + '_' + str(start_page) + '_' + str(end_page-1) + '.xlsx')
xl.close()
