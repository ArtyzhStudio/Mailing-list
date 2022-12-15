import smtplib
import openpyxl

import tkinter as tk
import tkinter.ttk as ttk

def logIn():
    global server
    server = smtplib.SMTP_SSL('smtp.gmail.com')
    server.set_debuglevel(1)
    server.login(log.get(), pswd.get())

def mail():
    wb = openpyxl.load_workbook(filename = '.\\table.xlsm')
    sheet = wb['Лист1']

    n = sheet.cell(1, 4).value
    m = sheet.max_row

    i = 1
    j = 1
    while i <= n and j <= m:
        if sheet.cell(i, 1).value != None:
            to = sheet.cell(i, 1).value
            head = sheet.cell(i, 2).value
            inner = sheet.cell(i, 3).value
            msg = f'''From: artyzhd@gmail.com
    Subject: {head}

    {inner}'''
            # print(head, msg)
            server.sendmail('arteck', to, msg)
            i += 1
        j += 1
    server.quit()

FILE = open('password.txt', 'r')
PASS = FILE.readline()
FILE.close()

win = tk.Tk()
win.title('Mail listing')
win.geometry('400x500')

login = ttk.Frame(win)
ttk.Label(login, text='Email login : ').pack(side='left')
log = ttk.Entry(login)
log.insert(0, 'artyzhd@gmail.com')
log.pack(side='right')
login.pack()

password = ttk.Frame(win)
ttk.Label(password, text='Email password : ').pack(side='left')
pswd = ttk.Entry(password)
pswd.insert(0, PASS)
pswd.pack(side='right')
password.pack()

buttons = ttk.Frame(win)
ttk.Button(buttons, text='Sign In', command=logIn).pack(side='left')
ttk.Button(buttons, text='Mail List', command=mail).pack(side='left')
buttons.pack()

print('ok')

server = smtplib.SMTP_SSL('smtp.gmail.com')
win.mainloop()