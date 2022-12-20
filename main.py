from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from os.path import basename

import json

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
    wb = openpyxl.load_workbook(filename = db.get())
    sheet = wb['Лист1']

    n = int(number.get())
    m = sheet.max_row
    i = 1
    j = 1
    while j <= n and i <= m:
        if sheet.cell(i, 1).value != None:
            to = sheet.cell(i, 1).value
            head = sheet.cell(i, 2).value
            inner = sheet.cell(i, 3).value
            files = sheet.cell(i, 4).value
            msg = MIMEMultipart()
            msg['From'] = log.get()
            msg['To'] = to
            msg['Subject'] = head
            msg.attach(MIMEText(inner))
            for t in files.split():
                with open(t, 'rb') as file:
                    add = MIMEApplication(file.read(), Name = basename(t))
                add['Content-Disposition'] = 'attachment; filename="%s"' % basename(t)
                msg.attach(add)
            server.sendmail('arteck', to, msg.as_string())
            j += 1
        i += 1
    server.quit()

FILE = open('password.txt', 'r')
PASS = FILE.readline()
FILE.close()

DBPATH = './table.xlsx'
FROM = 'artyzhd@gmail.com'

win = tk.Tk()
win.title('Mail listing')
win.geometry('400x500')

database = ttk.Frame(win)
ttk.Label(database, text='Table path : ').pack(side='left')
db = ttk.Entry(database)
db.insert(0, DBPATH)
db.pack(side='right')
database.pack(fill='x')

login = ttk.Frame(win)
ttk.Label(login, text='Email login : ').pack(side='left')
log = ttk.Entry(login)
log.insert(0, FROM)
log.pack(side='right')
login.pack(fill='x')

password = ttk.Frame(win)
ttk.Label(password, text='Email password : ').pack(side='left')
pswd = ttk.Entry(password)
pswd.insert(0, PASS)
pswd.pack(side='right')
password.pack(fill='x')

num = ttk.Frame(win)
ttk.Label(num, text='Number of mails : ').pack(side='left')
number = ttk.Entry(num)
number.insert(0, '5')
number.pack(side='right')
num.pack(fill='x')

buttons = ttk.Frame(win)
ttk.Button(buttons, text='Sign In', command=logIn).pack(side='left')
ttk.Button(buttons, text='Mail List', command=mail).pack(side='left')
buttons.pack(fill='x')

server = smtplib.SMTP_SSL('smtp.gmail.com')
win.mainloop()