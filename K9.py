from tkinter import *
import sqlite3
import hashlib
import math
import random
import smtplib
from tkinter import messagebox
from datetime import datetime
import subprocess
import win32com.client
from threading import Timer
from browser_history.browsers import Chrome
import win32evtlog
import os
from tkinter import filedialog
import winshell
import time
from cryptography.fernet import Fernet
import zipfile
from email.mime.base import MIMEBase
from email import encoders
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from PIL import Image, ImageTk
import re


with sqlite3.connect("k9DB.db") as db:
    cursor = db.cursor()

cursor.execute("""
CREATE TABLE IF NOT EXISTS caseinfo(
id INTEGER PRIMARY KEY,
email TEXT NOT NULL,
name TEXT NOT NULL, 
password TEXT NOT NULL,
remarks TEXT);
""")

cursor.execute("""
CREATE TABLE IF NOT EXISTS reporthash(
name TEXT PRIMARY KEY,
time TEXT NOT NULL, 
md5 TEXT NOT NULL,
pin TEXT);
""")

K9Root = Tk()
K9Root.geometry("1000x600+150+25")
K9Root.title("K9")
K9Root.iconbitmap("D:\Final Project\pythonProject\K-9 icon.ico")
K9Root.resizable(width=False, height=False)

Font = LARGE_FONT = ("Roboto", 12)
Font2 = ("Roboto", 10)


def hashPassword(input):
    hash = hashlib.md5(input)
    hash = hash.hexdigest()
    return hash

def NewCaseScreen():
    NewCase = Frame(K9Root, height='600', width='1000')
    frame = Frame(NewCase, height='450', width='750', highlightbackground='gray77', highlightthickness=2)
    NewCase.grid(row=1, column=0)
    frame.grid(padx=125, pady=75)
    NewCase.grid_propagate(False)
    frame.grid_propagate(False)

    def savedata():
        emailchar = '^[a-z0-9]+[\._]?[a-z0-9]+[@]\w+[.]\w{2,3}$'
        special_ch = ['~', '`', '!', '@', '#', '$', '%', '^', '&', '*', '(', ')', '-', '_', '+', '=', '{', '}', '[',
                      ']', '|', '\\', '/', ':', ';', '"', "'", '<', '>', ',', '.', '?']
        email = EmailEN.get()
        name = NameEN.get()
        remark = RemarkEN.get()
        pw = PWEN.get()
        rpw = REPWEN.get()

        try:
            if len(pw) == 0 & len(email) == 0 & len(name) == 0 & len(remark) == 0:
                lbl6N.config(text="Feilds are empty")
            elif not re.search(emailchar, email):
                lbl6N.config(text="Invalid Email")
            elif any(ch.isdigit() for ch in name):
                lbl6N.config(text="Name cannot contain numbers")
            elif not any(ch in special_ch for ch in pw):
                lbl6N.config(text="Password must contain a special character")
            elif not any(ch.isupper() for ch in pw):
                lbl6N.config(text="Password must contain an uppercase character")
            elif not any(ch.islower() for ch in pw):
                lbl6N.config(text="Password must contain a lowercase character")
            elif not any(ch.isdigit() for ch in pw):
                lbl6N.config(text="Password must contain a number")
            elif len(pw) < 8:
                lbl6N.config(text="Password must be not less than 8 characters")
            elif pw == rpw:
                hashedPassword = hashPassword(PWEN.get().encode('utf-8'))
                cursor.execute("DELETE FROM caseinfo WHERE id = 1")
                cursor.execute("DELETE FROM reporthash")
                save_data = """INSERT INTO caseinfo(email, name, password, remarks)
                VALUES(?,?,?,?)  """
                cursor.executemany(save_data, [(email, name, hashedPassword, remark)])
                db.commit()

                MainMenu()
            else:
                lbl6N.config(text="Passwords do not match")
        except Exception as e:
            messagebox.showinfo('Error', e)

    lbl7N = Label(frame, text="Create A New Case", font=Font)
    lbl1N = Label(frame, text="Email", font=Font2)
    EmailEN = Entry(frame, width=60, font=Font2, bg='gray77')
    EmailEN.focus()
    lbl2N = Label(frame, text="Investigator Name", font=Font2)
    NameEN = Entry(frame, width=60, font=Font2, bg='gray77')
    lbl3N = Label(frame, text="Password", font=Font2)
    PWEN = Entry(frame, width=60, show="*", font=Font2, bg='gray77')
    lbl4N = Label(frame, text="Re-Enter Password", font=Font2)
    REPWEN = Entry(frame, width=60, show="*", font=Font2, bg='gray77')
    lbl5N = Label(frame, text="Remarks", font=Font2)
    RemarkEN = Entry(frame, width=60, font=Font2, bg='gray77')
    lbl6N = Label(frame, font=Font)
    CreateBN = Button(frame, text="Create", command=savedata, height=2, width=20, font=Font2, bg='gray77', bd=1)

    lbl7N.grid(row=0, column=0, columnspan=2, padx=50, pady=10)
    lbl1N.grid(row=1, column=0, padx=(80, 30), pady=(40, 10), sticky='e')
    EmailEN.grid(row=1, column=1, pady=(40, 10))
    lbl2N.grid(row=2, column=0, padx=(80, 30), pady=10, sticky='e')
    NameEN.grid(row=2, column=1, pady=10)
    lbl3N.grid(row=3, column=0, padx=(80, 30), pady=10, sticky='e')
    PWEN.grid(row=3, column=1, pady=10)
    lbl4N.grid(row=4, column=0, padx=(80, 30), pady=10, sticky='e')
    REPWEN.grid(row=4, column=1, pady=10)
    lbl5N.grid(row=5, column=0, padx=(80, 30), pady=10, sticky='e')
    RemarkEN.grid(row=5, column=1, pady=10)
    lbl6N.grid(row=6, column=1, columnspan=2)
    CreateBN.grid(row=7, column=1, pady=10)



def loginScreen():
    Login = Frame(K9Root, height='600', width='1000')
    frame = Frame(Login, height='300', width='750', highlightbackground='gray77', highlightthickness=2)
    Login.grid()
    frame.grid(row=1, column=0, padx=125, pady=75)
    Login.grid_propagate(False)
    frame.grid_propagate(False)
    Lbl = Label(Login, text="WELCOME BACK!!!", font=Font)
    Lbl.grid(row=0, column=0, pady=(50, 0))

    def checkPassword():
        checkHashedPassword = hashPassword(PWEL.get().encode('utf-8'))
        cursor.execute("SELECT * FROM caseinfo WHERE id = 1 AND password = ?", [(checkHashedPassword)])
        getpw = cursor.fetchall()
        match = getpw
        if match:
            MainMenu()
        else:
            PWEL.delete(0, 'end')
            lbl2L.config(text="Wrong password")

    lbl3L = Label(frame, text="or", font=Font2)
    lbl1L = Label(frame, text="Password", font=Font2)
    PWEL = Entry(frame, width=60, font=Font2, bg='gray77', show="*")
    PWEL.focus()
    lbl2L = Label(frame, font=Font2)
    LoginBL = Button(frame, text="Login", command=checkPassword, font=Font2, height=2, width=20, bg='gray77', bd=1)

    lbl1L.grid(row=0, column=0, padx=(80, 50), pady=(80, 50))
    PWEL.grid(row=0, column=1, padx=0, pady=(80, 50), columnspan=3)
    lbl2L.grid(row=1, column=1, padx=0, pady=0)
    LoginBL.grid(row=2, column=1, padx=0, pady=0)
    lbl3L.grid(row=2, column=2)

    def otpscreen():
        for widget in K9Root.winfo_children():
            widget.destroy()

        OTPFrame = Frame(K9Root, height='600', width='1000')
        frame2 = Frame(OTPFrame, height='300', width='750', highlightbackground='gray77', highlightthickness=2)
        OTPFrame.grid()
        frame2.grid(row=1, column=0, padx=125, pady=125)
        OTPFrame.grid_propagate(False)
        frame2.grid_propagate(False)

        otplbl3 = Label(frame2, text="Please Enter Your OTP Code", font=Font)
        otplbl3.grid(row=0, column=0, columnspan=2, pady=(50, 10))
        otplbl1 = Label(frame2, text="Email", font=Font2)
        otptxt1 = Entry(frame2, width=60, font=Font2, bg='gray77', bd=1)
        cursor.execute("SELECT email FROM caseinfo WHERE id = 1")
        otptxt1.delete(0, END)
        x = cursor.fetchall()
        for row in x:
            otptxt1.insert(END, row)
        otptxt1.configure(state=DISABLED)
        otplbl2 = Label(frame2, text="Enter OTP code", font=Font2)
        otptxt2 = Entry(frame2, width=60, show="*", font=Font2, bg='gray77', bd=1)
        otptxt2.focus()

        otplbl1.grid(row=1, column=00, padx=(100, 50), pady=(30, 20), sticky='e')
        otptxt1.grid(row=1, column=1, padx=0, pady=(30, 20))
        otplbl2.grid(row=2, column=0, padx=(100, 50), sticky='e')
        otptxt2.grid(row=2, column=1, pady=20)

        digits = "0123456789"
        OTP = ""
        for i in range(6):
            OTP += digits[math.floor(random.random() * 10)]
        otp = OTP + " is your OTP"

        s = smtplib.SMTP('smtp.gmail.com', 587)
        s.starttls()
        s.login("k9.official.app@gmail.com", "sxpfqvhnkjrqvpbm")
        subject = "Your OTP code"
        msg = MIMEText(otp)
        msg['Subject'] = subject
        cursor.execute("SELECT * from caseinfo")
        emailid = str(cursor.fetchall()[0][1])
        s.sendmail('&&&&&&&&&&&', emailid, msg.as_string())

        def verifyOTP():
            otpcode = otptxt2.get()
            try:
                int(otpcode)
                if int(otpcode) == int(OTP):
                    resetpwscreen()
                else:
                    messagebox.showinfo('Wrong OTP Code', 'We have sent another OTP code to your email!')
                    otptxt2.delete(0, END)
                    otpscreen()
            except Exception:
                messagebox.showinfo('Wrong OTP Code', 'We have sent another OTP code to your email!')
                otpscreen()

        def resetpwscreen():
            for widget in K9Root.winfo_children():
                widget.destroy()

            ResetFrame = Frame(K9Root, height='600', width='1000')
            frame3 = Frame(ResetFrame, height='300', width='750', highlightbackground='gray77', highlightthickness=2)
            ResetFrame.grid()
            frame3.grid(row=1, column=0, padx=125, pady=125)
            ResetFrame.grid_propagate(False)
            frame3.grid_propagate(False)

            rlbl3 = Label(frame3, text="Reset Your Password", font=Font)
            rlbl3.grid(row=0, column=0, columnspan=2, pady=(50, 10))
            rlbl1 = Label(frame3, text="Password", font=Font2)
            rtxt1 = Entry(frame3, width=60, show="*", font=Font2, bg='gray77', bd=1)
            rtxt1.focus()
            rlbl2 = Label(frame3, text="Enter OTP code", font=Font2)
            rtxt2 = Entry(frame3, width=60, show="*", font=Font2, bg='gray77', bd=1)
            rlbl4 = Label(frame3, font=Font2)

            rlbl1.grid(row=1, column=00, padx=(100, 50), pady=(30, 20), sticky='e')
            rtxt1.grid(row=1, column=1, padx=0, pady=(30, 20))
            rlbl2.grid(row=2, column=0, padx=(100, 50), sticky='e')
            rtxt2.grid(row=2, column=1, pady=20)
            rlbl4.grid(row=3, column=1)

            def reset():
                special_ch = ['~', '`', '!', '@', '#', '$', '%', '^', '&', '*', '(', ')', '-', '_', '+', '=', '{', '}',
                              '[',']', '|', '\\', '/', ':', ';', '"', "'", '<', '>', ',', '.', '?']
                pw = rtxt1.get()
                rpw = rtxt2.get()
                try:
                    if len(pw) == 0 or len(rpw) == 0:
                        rlbl4.config(text="Fields are empty")
                    elif not any(ch.isdigit() for ch in pw):
                        rlbl4.config(text="Password must contain a number")
                    elif not any(ch in special_ch for ch in pw):
                        rlbl4.config(text="Password must contain a special character")
                    elif not any(ch.isupper() for ch in pw):
                        rlbl4.config(text="Password must contain an uppercase character")
                    elif not any(ch.islower() for ch in pw):
                        rlbl4.config(text="Password must contain a lowercase character")
                    elif not any(ch.isdigit() for ch in pw):
                        rlbl4.config(text="Password must contain a number")
                    elif len(pw) < 8:
                        rlbl4.config(text="Password must be not less than 8 characters")
                    elif pw == rpw:
                        hashedPassword = hashPassword(pw.encode('utf-8'))
                        update_password = ("""UPDATE caseinfo SET
                        password = ?
                        WHERE id = 1 """)
                        cursor.execute(update_password, [(hashedPassword)])
                        db.commit()
                        for widget in K9Root.winfo_children():
                            widget.destroy()
                        loginScreen()
                    else:
                        rlbl4.config(text="Passwords do not match")
                except Exception as e:
                    rlbl4.config(text=e)

            rbtn = Button(frame3, text="Reset", command=reset, font=Font2, height=2, width=20, bg='gray77', bd=1)
            rbtn.grid(row=4, column=1)

        otpbtn = Button(frame2, text="Next", command=verifyOTP, font=Font2, height=2, width=20, bg='gray77', bd=1)
        otpbtn.grid(row=3, column=1)

    fpbtn = Button(frame, text="Forget Password?", command=otpscreen, font=Font2, height=2, width=20, bd=0)
    fpbtn.grid(row=2, column=0, padx=50)

    def GoNewCase():
        for widget in K9Root.winfo_children():
            widget.destroy()
        NewCaseScreen()

    ncbtn = Button(frame, text="New Case", command=GoNewCase, font=Font2, height=2, width=20, bg='gray77', bd=1)
    ncbtn.grid(row=2, column=3)


def md5sum(filename, blocksize=65536):
    try:
        hash = hashlib.md5()
        with open(filename, "rb") as f:
            for block in iter(lambda: f.read(blocksize), b""):
                hash.update(block)
        reportHash = hash.hexdigest()
        filename = os.path.basename(filename)
        save_data = """INSERT OR REPLACE INTO reporthash(name, time, md5)
                    VALUES(?, ?, ?)"""
        cursor.executemany(save_data, [(filename, str(datetime.now()), reportHash)])
        db.commit()
    except Exception as e:
        messagebox.showinfo('Error', e)

def pininlog(filename, pin):
    try:
        save_data = """UPDATE reporthash SET pin = ? WHERE name = ? """
        cursor.execute(save_data, (pin, filename))
        db.commit()
    except Exception as e:
        messagebox.showinfo('Error', e)

def removepin(filename):
    try:
        pin = None
        remove_data = """UPDATE reporthash SET pin = ? WHERE name = ? """
        cursor.execute(remove_data, (pin, filename))
        db.commit()
    except Exception as e:
        messagebox.showinfo('Error', e)

path = "Reports"
isExist = os.path.exists(path)
if not isExist:
    try:
        os.makedirs(path)
    except Exception as e:
        messagebox.showinfo('Error', e)

def MainMenu():
    for widget in K9Root.winfo_children():
        widget.destroy()

    Mainmenu = Frame(K9Root, height=600, width=1000)
    Mainmenu.grid()

    NavBar = Frame(Mainmenu, height=600, width=200)
    # NavBar.grid_propagate(False)
    NavBar.grid(row=0, column=0)

    Tabs = Frame(Mainmenu, height=600, width=800)
    Tabs.grid(row=0, column=1)

    def GoHome():
        for widget in Tabs.winfo_children():
           widget.destroy()
        HomeF = Frame(Tabs, height=600, width=800)
        HomeF.grid(row=0, column=0)
        HomeF.grid_propagate(False)
        HomeF.configure(bg='white')

        global img
        path = r"K-9logo.png"
        img = ImageTk.PhotoImage(Image.open(path))
        panel = Label(HomeF, image=img, bd=0)
        panel.grid(padx=100)

    GoHome()

    def GoDevice():
        for widget in Tabs.winfo_children():
            widget.destroy()
        navbarbtns = [devicebtn, appbtn, windbtn, multbtn, logbtn, encryptbtn, sendbtn]
        for x in navbarbtns:
            x.configure(bg='gray77')
        devicebtn.configure(bg='#885933')
        DeviceF = Frame(Tabs, height=600, width=800)
        DeviceF.grid(row=0, column=0)
        DeviceF.grid_propagate(False)

        def GoDrpF():
            DrpF = Frame(DeviceF, height=540, width=800)
            DrpF.grid(row=1, column=0, columnspan=8)
            DrpF.grid_propagate(False)

            Dnavbarbtns = [Dbtn1, Dbtn2, Dbtn3, Dbtn4, Dbtn5]
            for x in Dnavbarbtns:
                x.configure(bg='#D4BEAD')
            Dbtn1.configure(bg='#A07858')

            def tasklist():
                try:
                    response = os.popen('tasklist')
                    Drptxt.configure(state=NORMAL)
                    Drptxt.insert(END, "\nTask List:\n")
                    for x in response:
                        Drptxt.insert(END, x)
                    Drptxt.configure(state=DISABLED)
                except Exception as e:
                    messagebox.showinfo('Error', e)

            def physicalmemory():
                try:
                    response = os.popen('systeminfo |find "Available Physical Memory"')
                    Drptxt.configure(state=NORMAL)
                    Drptxt.insert(END, "\nPhysical Memory:\n")
                    for x in response:
                        Drptxt.insert(END, x)
                    Drptxt.configure(state=DISABLED)
                except Exception as e:
                    messagebox.showinfo('Error', e)

            def cpudetails():
                try:
                    response = os.popen('wmic cpu get caption, deviceid, name, numberofcores, maxclockspeed, status')
                    Drptxt.configure(state=NORMAL)
                    Drptxt.insert(END, "\nCPU Details:\n")
                    for x in response:
                        Drptxt.insert(END, x)
                    Drptxt.configure(state=DISABLED)
                except Exception as e:
                    messagebox.showinfo('Error', e)

            def Drpexport():
                try:
                    DrpReport = open('Reports/Device_Configurations.txt', 'a')
                    DrpReport.write("Time exported : " + str(datetime.now()) +
                                    "\n\n__RUNNING PROCESSES__\n\n")
                    response = Drptxt.get("1.0", END)
                    for x in response:
                        DrpReport.write(x)
                    DrpReport.close()
                    path = os.path.dirname(os.path.realpath('Reports/Device_Configurations.txt'))
                    messagebox.showinfo('Export Complete',
                                        'Your records has been successfully exported to: \n' + path + '\\' + 'Device_Configurations.txt')
                    md5sum("Reports/Device_Configurations.txt")
                except Exception as e:
                    messagebox.showinfo('Error', e)

            Drpbtn1 = Button(DrpF, text='Task List', height=2, width=22, command=tasklist, bg='gray77')
            Drpbtn2 = Button(DrpF, text='Available Physical\nMemory', height=2, width=22, command=physicalmemory, bg='gray77')
            Drpbtn3 = Button(DrpF, text='CPU Details', height=2, width=22, command=cpudetails, bg='gray77')
            Drpbtn4 = Button(DrpF, text='Export', height=2, width=22, command=Drpexport, bg='gray77')

            Drpscrolly = Scrollbar(DrpF, orient=VERTICAL)
            Drpscrollx =Scrollbar(DrpF, orient=HORIZONTAL)
            Drptxt = Text(DrpF, yscrollcommand=Drpscrolly.set, xscrollcommand=Drpscrollx.set, height=22, width=90, wrap='none')
            Drpscrolly.configure(command=Drptxt.yview)
            Drpscrollx.configure(command=Drptxt.xview)
            Drptxt.configure(state=DISABLED)

            Drpbtn1.grid(row=2, column=0, padx=30, pady=20, sticky='ew')
            Drpbtn2.grid(row=2, column=1, padx=30, pady=20, sticky='ew')
            Drpbtn3.grid(row=2, column=2, padx=30, pady=20, sticky='ew')
            Drptxt.grid(row=3, column=0, columnspan=3, sticky='e', padx=(30, 0))
            Drpscrolly.grid(row=3, column=3, sticky='nsw')
            Drpscrollx.grid(row=4, column=0, columnspan=3, sticky='ew', padx=(30, 0))
            Drpbtn4.grid(row=5, column=2, padx=30, pady=20, sticky='ew')

        def GoDnsF():
            DnsF = Frame(DeviceF, height=540, width=800)
            DnsF.grid(row=1, column=0, columnspan=8)
            DnsF.grid_propagate(False)

            Dnavbarbtns = [Dbtn1, Dbtn2, Dbtn3, Dbtn4, Dbtn5]
            for x in Dnavbarbtns:
                x.configure(bg='#D4BEAD')
            Dbtn2.configure(bg='#A07858')

            def NetConfig():
                try:
                    response = os.popen('ipconfig /all')
                    Dnstxt.configure(state=NORMAL)
                    Dnstxt.insert(END, "\nAll Network Configurations:\n")
                    for x in response:
                        Dnstxt.insert(END, x)
                    Dnstxt.configure(state=DISABLED)
                except Exception as e:
                    messagebox.showinfo('Error', e)

            def TCPIPCon():
                try:
                    response = subprocess.Popen('netstat', stdout=subprocess.PIPE)
                    timeout_sec = 2
                    timer = Timer(timeout_sec, response.kill)
                    timer.start()
                    response.stdout = response.communicate()
                    timer.cancel()
                    Dnstxt.configure(state=NORMAL)
                    Dnstxt.insert(END, "\nTCP/IP Connections for 2 seconds:\n")
                    for x in response.stdout:
                        Dnstxt.insert(END, x)
                    Dnstxt.configure(state=DISABLED)
                except TclError:
                    pass
                except Exception as e:
                    messagebox.showinfo('Error', e)

            def ActiveUsedIP():
                try:
                    response = os.popen('arp -a')
                    Dnstxt.configure(state=NORMAL)
                    Dnstxt.insert(END, "\nActive and Used IP Addresses:\n")
                    for x in response:
                        Dnstxt.insert(END, x)
                    Dnstxt.configure(state=DISABLED)
                except Exception as e:
                    messagebox.showinfo('Error', e)

            def ping():
                try:
                    address = Dnsentry.get()
                    response = os.popen('ping ' + address)
                    Dnstxt.configure(state=NORMAL)
                    Dnstxt.insert(END, "\n\nPing Results of : " + address + "\n\n")
                    for x in response:
                        Dnstxt.insert(END, x)
                    Dnstxt.configure(state=DISABLED)
                except ValueError:
                    messagebox.showinfo('Error', ValueError)

            def tracert():
                try:
                    address = Dnsentry.get()
                    response = os.popen('tracert -h 2 ' + address)
                    Dnstxt.configure(state=NORMAL)
                    Dnstxt.insert(END, "\n\ntrace route Results of : " + address + "\n\n")
                    for x in response:
                        Dnstxt.insert(END, x)
                    Dnstxt.configure(state=DISABLED)
                except Exception as e:
                    messagebox.showinfo('Error', e)

            def Dnsexport():
                try:
                    DnsReport = open('Reports/Device_Configurations.txt', 'a')
                    DnsReport.write("Time exported : " + str(datetime.now()) +
                                    "\n\n__NETWORK STATICS__\n\n")
                    response = Dnstxt.get("1.0", END)
                    for x in response:
                        DnsReport.write(x)
                    DnsReport.close()
                    path = os.path.dirname(os.path.realpath('Reports/Device_Configurations.txt'))
                    messagebox.showinfo('Export Complete',
                                        'Your records has been successfully exported to: \n' + path + '\\' + 'Device_Configurations.txt')
                    md5sum("Reports/Device_Configurations.txt")
                except Exception as e:
                    messagebox.showinfo('Error', e)

            Dnsbtn1 = Button(DnsF, text='All Network Configurations', height=2, width=15, command=NetConfig, bg='gray77')
            Dnsbtn2 = Button(DnsF, text='TCP/IP Connections', height=2, width=15, command=TCPIPCon, bg='gray77')
            Dnsbtn3 = Button(DnsF, text='Active & Used\nIP Addresses', height=2, width=15, command=ActiveUsedIP, bg='gray77')
            Dnsbtn4 = Button(DnsF, text='Ping', height=2, width=5, command=ping, bg='gray77')
            Dnsbtn5 = Button(DnsF, text='Trace Route', height=2, width=5, command=tracert, bg='gray77')
            Dnsbtn6 = Button(DnsF, text='Export', height=2, width=12, command=Dnsexport, bg='gray77')

            Dnsentry = Entry(DnsF, width=60, font=Font2)
            Dnsscrolly = Scrollbar(DnsF, orient=VERTICAL)
            Dnsscrollx = Scrollbar(DnsF, orient=HORIZONTAL)
            Dnstxt = Text(DnsF, yscrollcommand=Dnsscrolly.set, xscrollcommand=Dnsscrollx.set, height=20, width=90, wrap='none')
            Dnsscrolly.configure(command=Dnstxt.yview)
            Dnsscrollx.configure(command=Dnstxt.xview)
            Dnstxt.configure(state=DISABLED)

            Dnsbtn1.grid(row=2, column=0, columnspan=2, padx=30, pady=10, sticky='ew')
            Dnsbtn2.grid(row=2, column=2, padx=30, pady=10, sticky='ew')
            Dnsbtn3.grid(row=2, column=3, padx=30, pady=10, sticky='ew')
            Dnsbtn4.grid(row=3, column=0, padx=10, pady=10, sticky='ew')
            Dnsbtn5.grid(row=3, column=1, padx=10, pady=10, sticky='ew')
            Dnsentry.grid(row=3, column=2, columnspan=2)
            Dnstxt.grid(row=4, column=0, padx=(30, 0), columnspan=4, sticky='e')
            Dnsscrolly.grid(row=4, column=4, sticky='nsw')
            Dnsscrollx.grid(row=5, column=0, columnspan=4, padx=(30, 0), sticky='ew')
            Dnsbtn6.grid(row=6, column=3, pady=10, columnspan=3)

        def GoDsysiF():
            DsysiF = Frame(DeviceF, height=540, width=800)
            DsysiF.grid(row=1, column=0, columnspan=8)
            DsysiF.grid_propagate(False)

            Dnavbarbtns = [Dbtn1, Dbtn2, Dbtn3, Dbtn4, Dbtn5]
            for x in Dnavbarbtns:
                x.configure(bg='#D4BEAD')
            Dbtn3.configure(bg='#A07858')

            def ovsysinfo():
                try:
                    response = os.popen('systeminfo')
                    Dsystxt.configure(state=NORMAL)
                    Dsystxt.insert(END, "\n\nOverall System Info : \n\n")
                    for x in response:
                        Dsystxt.insert(END, x)
                    Dsystxt.configure(state=DISABLED)
                except Exception as e:
                    messagebox.showinfo('Error', e)

            def useracc():
                try:
                    response = os.popen('wmic USERACCOUNT')
                    Dsystxt.configure(state=NORMAL)
                    Dsystxt.insert(END, "\n\nUser Accounts: \n\n")
                    for x in response:
                        Dsystxt.insert(END, x)
                    Dsystxt.configure(state=DISABLED)
                except Exception as e:
                    messagebox.showinfo('Error', e)

            def nicinfo():
                try:
                    response = os.popen('wmic NIC')
                    Dsystxt.configure(state=NORMAL)
                    Dsystxt.insert(END, "\n\nNIC Info : \n\n")
                    for x in response:
                        Dsystxt.insert(END, x)
                    Dsystxt.configure(state=DISABLED)
                except Exception as e:
                    messagebox.showinfo('Error', e)

            def raminfo():
                try:
                    response = os.popen('wmic MEMORYCHIP')
                    Dsystxt.configure(state=NORMAL)
                    Dsystxt.insert(END, "\n\nRAM Info : \n\n")
                    for x in response:
                        Dsystxt.insert(END, x)
                    Dsystxt.configure(state=DISABLED)
                except Exception as e:
                    messagebox.showinfo('Error', e)

            def cpuinfo():
                try:
                    response = os.popen('wmic CPU')
                    Dsystxt.configure(state=NORMAL)
                    Dsystxt.insert(END, "\n\nCPU Info : \n\n")
                    for x in response:
                        Dsystxt.insert(END, x)
                    Dsystxt.configure(state=DISABLED)
                except Exception as e:
                    messagebox.showinfo('Error', e)

            def biosinfo():
                try:
                    response = os.popen('wmic BIOS')
                    Dsystxt.configure(state=NORMAL)
                    Dsystxt.insert(END, "\n\nBIOS Info : \n\n")
                    for x in response:
                        Dsystxt.insert(END, x)
                    Dsystxt.configure(state=DISABLED)
                except Exception as e:
                    messagebox.showinfo('Error', e)

            def osinfo():
                try:
                    response = os.popen('wmic OS')
                    Dsystxt.configure(state=NORMAL)
                    Dsystxt.insert(END, "\n\nOS Info : \n\n")
                    for x in response:
                        Dsystxt.insert(END, x)
                    Dsystxt.configure(state=DISABLED)
                except Exception as e:
                    messagebox.showinfo('Error', e)

            def motherboardinfo():
                try:
                    response = os.popen('wmic BASEBOARD')
                    Dsystxt.configure(state=NORMAL)
                    Dsystxt.insert(END, "\n\nMotherboard Info : \n\n")
                    for x in response:
                        Dsystxt.insert(END, x)
                    Dsystxt.configure(state=DISABLED)
                except Exception as e:
                    messagebox.showinfo('Error', e)

            def Dsysexport():
                try:
                    DsysReport = open('Reports/Device_Configurations.txt', 'a')
                    DsysReport.write("Time exported : " + str(datetime.now()) +
                                     "\n\n__SYSTEM INFO AND CONFIGURATIONS__\n\n")
                    response = Dsystxt.get("1.0", END)
                    for x in response:
                        DsysReport.write(x)
                    DsysReport.close()
                    path = os.path.dirname(os.path.realpath('Reports/Device_Configurations.txt'))
                    messagebox.showinfo('Export Complete',
                                        'Your records has been successfully exported to: \n' + path + '\\' + 'Device_Configurations.txt')
                    md5sum("Reports/Device_Configurations.txt")
                except Exception as e:
                    messagebox.showinfo('Error', e)

            Dsysbtn1 = Button(DsysiF, text='Overall System\nInfo', height=2, width=18, command=ovsysinfo, bg='gray77')
            Dsysbtn2 = Button(DsysiF, text='User Account\nInfo', height=2, width=18, command=useracc, bg='gray77')
            Dsysbtn3 = Button(DsysiF, text='NIC Info\n(MAC)', height=2, width=18, command=nicinfo, bg='gray77')
            Dsysbtn4 = Button(DsysiF, text='RAM Info', height=2, width=18, command=raminfo, bg='gray77')
            Dsysbtn5 = Button(DsysiF, text='CPU Info', height=2, width=18, command=cpuinfo, bg='gray77')
            Dsysbtn6 = Button(DsysiF, text='BIOS Info', height=2, width=18, command=biosinfo, bg='gray77')
            Dsysbtn7 = Button(DsysiF, text='OS Info', height=2, width=18, command=osinfo, bg='gray77')
            Dsysbtn8 = Button(DsysiF, text='Motherboard Info', height=2, width=18, command=motherboardinfo, bg='gray77')
            Dsysbtn9 = Button(DsysiF, text='Export', height=2, width=18, command=Dsysexport, bg='gray77')

            Dsysscrolly = Scrollbar(DsysiF, orient=VERTICAL)
            Dsysscrollx = Scrollbar(DsysiF, orient=HORIZONTAL)
            Dsystxt = Text(DsysiF, yscrollcommand=Dsysscrolly.set, xscrollcommand=Dsysscrollx.set, height=22, width=90, wrap='none')
            Dsysscrolly.configure(command=Dsystxt.yview)
            Dsysscrollx.configure(command=Dsystxt.xview)
            Dsystxt.configure(state=DISABLED)

            Dsysbtn1.grid(row=1, column=0, padx=(30, 5), pady=5, sticky='ew')
            Dsysbtn2.grid(row=1, column=1, padx=5, pady=5, sticky='ew')
            Dsysbtn3.grid(row=1, column=2, padx=5, pady=5, sticky='ew')
            Dsysbtn4.grid(row=1, column=3, padx=5, pady=5, sticky='ew')
            Dsysbtn5.grid(row=2, column=0, padx=(30, 5), pady=5, sticky='ew')
            Dsysbtn6.grid(row=2, column=1, padx=5, pady=5, sticky='ew')
            Dsysbtn7.grid(row=2, column=2, padx=5, pady=5, sticky='ew')
            Dsysbtn8.grid(row=2, column=3, padx=5, pady=5, sticky='ew')
            Dsystxt.grid(row=3, column=0, columnspan=4, sticky='we', padx=(30, 0))
            Dsysscrolly.grid(row=3, column=4, sticky='nsw')
            Dsysscrollx.grid(row=4, column=0, columnspan=4, sticky='we', padx=(30, 0))
            Dsysbtn9.grid(row=5, column=3, padx=10, pady=20, sticky='ew')

        def GoDstoiF():
            DstoiF = Frame(DeviceF, height=540, width=800)
            DstoiF.grid(row=1, column=0, columnspan=8)
            DstoiF.grid_propagate(False)

            Dnavbarbtns = [Dbtn1, Dbtn2, Dbtn3, Dbtn4, Dbtn5]
            for x in Dnavbarbtns:
                x.configure(bg='#D4BEAD')
            Dbtn4.configure(bg='#A07858')

            def diskinfo():
                try:
                    response = os.popen('wmic DISKDRIVE')
                    Dstotxt.configure(state=NORMAL)
                    Dstotxt.insert(END, "\n\nDisk Drive Info : \n\n")
                    for x in response:
                        Dstotxt.insert(END, x)
                    Dstotxt.configure(state=DISABLED)
                except Exception as e:
                    messagebox.showinfo('Error', e)

            def diskpartition():
                try:
                    response = os.popen('wmic PARTITION')
                    Dstotxt.configure(state=NORMAL)
                    Dstotxt.insert(END, "\n\nDisk Partition Info : \n\n")
                    for x in response:
                        Dstotxt.insert(END, x)
                    Dstotxt.configure(state=DISABLED)
                except Exception as e:
                    messagebox.showinfo('Error', e)

            def logicaldiskinfo():
                try:
                    response = os.popen('wmic LOGICALDISK')
                    Dstotxt.configure(state=NORMAL)
                    Dstotxt.insert(END, "\n\nLogical Disk Info : \n\n")
                    for x in response:
                        Dstotxt.insert(END, x)
                    Dstotxt.configure(state=DISABLED)
                except Exception as e:
                    messagebox.showinfo('Error', e)

            def clusterinfo():
                try:
                    response = os.popen('fsutil fsinfo ntfsinfo c:')
                    Dstotxt.configure(state=NORMAL)
                    Dstotxt.insert(END, "\n\nCluster Sector Info : \n\n")
                    for x in response:
                        Dstotxt.insert(END, x)
                    Dstotxt.configure(state=DISABLED)
                except Exception as e:
                    messagebox.showinfo('Error', e)

            def Dstoexport():
                try:
                    DstoReport = open('Reports/Device_Configurations.txt', 'a')
                    DstoReport.write("Time exported : " + str(datetime.now()) +
                                     "\n\n__SYSTEM INFO AND CONFIGURATIONS__\n\n")
                    response = Dstotxt.get("1.0", END)
                    for x in response:
                        DstoReport.write(x)
                    DstoReport.close()
                    path = os.path.dirname(os.path.realpath('Reports/Device_Configurations.txt'))
                    messagebox.showinfo('Export Complete',
                                        'Your records has been successfully exported to: \n' + path + '\\' + 'Device_Configurations.txt')
                    md5sum("Reports/Device_Configurations.txt")
                except Exception as e:
                    messagebox.showinfo('Error', e)

            Dstobtn1 = Button(DstoiF, text='Disk Info\n(Physical Drive)', height=2, width=30, command=diskinfo, bg='gray77')
            Dstobtn2 = Button(DstoiF, text='Disk Partition\nInfo', height=2, width=30, command=diskpartition, bg='gray77')
            Dstobtn3 = Button(DstoiF, text='Logical Disk\nInfo', height=2, width=30, command=logicaldiskinfo, bg='gray77')
            Dstobtn4 = Button(DstoiF, text='Cluster/Sector\nInfo', height=2, width=30, command=clusterinfo, bg='gray77')
            Dstobtn5 = Button(DstoiF, text='Export', height=2, width=12, command=Dstoexport, bg='gray77')

            Dstoscrolly = Scrollbar(DstoiF, orient=VERTICAL)
            Dstoscrollx = Scrollbar(DstoiF, orient=HORIZONTAL)
            Dstotxt = Text(DstoiF, yscrollcommand=Dstoscrolly.set, xscrollcommand=Dstoscrollx.set, height=22, wrap='none')
            Dstoscrolly.configure(command=Dstotxt.yview)
            Dstoscrollx.configure(command=Dstotxt.xview)
            Dstotxt.configure(state=DISABLED)

            Dstobtn1.grid(row=1, column=0, padx=20, pady=5, sticky='ew')
            Dstobtn2.grid(row=1, column=1, padx=20, pady=5, sticky='ew')
            Dstobtn3.grid(row=1, column=2, padx=20, pady=5, sticky='ew')
            Dstobtn4.grid(row=2, column=1, padx=20, pady=5, sticky='ew')
            Dstotxt.grid(row=3, column=0, columnspan=3, sticky='we', padx=(20, 0))
            Dstoscrolly.grid(row=3, column=3, sticky='nsw')
            Dstoscrollx.grid(row=4, column=0, columnspan=3, sticky='we', padx=(20, 0))
            Dstobtn5.grid(row=5, column=2, padx=10, pady=20, sticky='ew')

        def GoDpnpF():
            DpnpF = Frame(DeviceF, height=540, width=800)
            DpnpF.grid(row=1, column=0, columnspan=8)
            DpnpF.grid_propagate(False)

            Dnavbarbtns = [Dbtn1, Dbtn2, Dbtn3, Dbtn4, Dbtn5]
            for x in Dnavbarbtns:
                x.configure(bg='#D4BEAD')
            Dbtn5.configure(bg='#A07858')

            def allpnp():
                try:
                    response = subprocess.Popen(['powershell.exe', 'Get-PnpDevice'], stdout=subprocess.PIPE)
                    Dpnptxt.configure(state=NORMAL)
                    Dpnptxt.insert(END, "\n\nAll PnP Device List: \n\n")
                    for x in response.stdout:
                        Dpnptxt.insert(END, x)
                    Dpnptxt.configure(state=DISABLED)
                except Exception as e:
                    messagebox.showinfo('Error', e)

            def usbonly():
                try:
                    response = subprocess.Popen(['powershell.exe', "Get-PnpDevice -Class 'USB'"], stdout=subprocess.PIPE)
                    Dpnptxt.configure(state=NORMAL)
                    Dpnptxt.insert(END, "\n\nPnP USB Only Device List: \n\n")
                    for x in response.stdout:
                        Dpnptxt.insert(END, x)
                    Dpnptxt.configure(state=DISABLED)
                except Exception as e:
                    messagebox.showinfo('Error', e)

            def currentremovable():
                try:
                    response = os.popen('wmic logicaldisk where "drivetype =2"')
                    Dpnptxt.configure(state=NORMAL)
                    Dpnptxt.insert(END, "\n\nCurrent Removable Device List: \n\n")
                    for x in response:
                        Dpnptxt.insert(END, x)
                    Dpnptxt.configure(state=DISABLED)
                except Exception as e:
                    messagebox.showinfo('Error', e)

            def allremovable():
                try:
                    run = "Get-ItemProperty -Path HKLM:/SYSTEM/CurrentControlSet/Enum/USBSTOR/*/* | Select FriendlyName"
                    response = subprocess.Popen(['powershell.exe', run], stdout=subprocess.PIPE)
                    Dpnptxt.configure(state=NORMAL)
                    Dpnptxt.insert(END, "\n\nAll Removable Device List: \n\n")
                    for x in response.stdout:
                        Dpnptxt.insert(END, x)
                    Dpnptxt.configure(state=DISABLED)
                except Exception as e:
                    messagebox.showinfo('Error', e)

            def Dpnpexport():
                try:
                    DpnpReport = open('Reports/Device_Configurations.txt', 'a')
                    DpnpReport.write("Time exported : " + str(datetime.now()) +
                                     "\n\n__SYSTEM INFO AND CONFIGURATIONS__\n\n")
                    response = Dpnptxt.get("1.0", END)
                    for x in response:
                        DpnpReport.write(x)
                    DpnpReport.close()
                    path = os.path.dirname(os.path.realpath('Reports/Device_Configurations.txt'))
                    messagebox.showinfo('Export Complete',
                                        'Your records has been successfully exported to: \n' + path + '\\' + 'Device_Configurations.txt')
                    md5sum("Reports/Device_Configurations.txt")
                except Exception as e:
                    messagebox.showinfo('Error', e)

            Dpnpbtn1 = Button(DpnpF, text='All PnP Device List', height=2, width=20, command=allpnp, bg='gray77')
            Dpnpbtn2 = Button(DpnpF, text='USB Devices Only', height=2, width=20, command=usbonly, bg='gray77')
            Dpnpbtn3 = Button(DpnpF, text='Current USB Removable\nDevice List', height=2, width=20,
                              command=currentremovable, bg='gray77')
            Dpnpbtn4 = Button(DpnpF, text='All USB Removable\nDevice List', height=2, width=20, command=allremovable, bg='gray77')
            Dpnpbtn5 = Button(DpnpF, text='Export', height=2, width=12, command=Dpnpexport, bg='gray77')

            Dpnpscrolly = Scrollbar(DpnpF, orient=VERTICAL)
            Dpnpscrollx = Scrollbar(DpnpF, orient=HORIZONTAL)
            Dpnptxt = Text(DpnpF, yscrollcommand=Dpnpscrolly.set, xscrollcommand=Dpnpscrollx.set, height=22, wrap='none')
            Dpnpscrolly.configure(command=Dpnptxt.yview)
            Dpnpscrollx.configure(command=Dpnptxt.xview)
            Dpnptxt.configure(state=DISABLED)

            Dpnpbtn1.grid(row=1, column=0, padx=20, pady=20, sticky='ew')
            Dpnpbtn2.grid(row=1, column=1, padx=20, pady=20, sticky='ew')
            Dpnpbtn3.grid(row=1, column=2, padx=20, pady=20, sticky='ew')
            Dpnpbtn4.grid(row=1, column=3, padx=20, pady=20, sticky='ew')
            Dpnptxt.grid(row=3, column=0, columnspan=4, sticky='we', padx=(20, 0))
            Dpnpscrolly.grid(row=3, column=4, sticky='nsw')
            Dpnpscrollx.grid(row=4, column=0, columnspan=4, sticky='we', padx=(20, 0))
            Dpnpbtn5.grid(row=5, column=3, padx=10, pady=20, sticky='ew')

        Dbtn1 = Button(DeviceF, text='Running Processes', height=3, width=22, command=GoDrpF, bg='#D4BEAD')
        Dbtn2 = Button(DeviceF, text='Network Stats', height=3, width=22, command=GoDnsF, bg='#D4BEAD')
        Dbtn3 = Button(DeviceF, text='System Info', height=3, width=21, command=GoDsysiF, bg='#D4BEAD')
        Dbtn4 = Button(DeviceF, text='Storage Info', height=3, width=21, command=GoDstoiF, bg='#D4BEAD')
        Dbtn5 = Button(DeviceF, text='PnP Devices', height=3, width=21, command=GoDpnpF, bg='#D4BEAD')

        Dbtn1.grid(row=0, column=0, sticky='ew')
        Dbtn2.grid(row=0, column=1, sticky='ew')
        Dbtn3.grid(row=0, column=2, sticky='ew')
        Dbtn4.grid(row=0, column=3, sticky='ew')
        Dbtn5.grid(row=0, column=4, sticky='ew')

        GoDrpF()

    def GoApp():
        for widget in Tabs.winfo_children():
            widget.destroy()
        navbarbtns = [devicebtn, appbtn, windbtn, multbtn, logbtn, encryptbtn, sendbtn]
        for x in navbarbtns:
            x.configure(bg='gray77')
        appbtn.configure(bg='#885933')
        AppF = Frame(Tabs, height=600, width=800)
        AppF.grid(row=0, column=0)
        AppF.grid_propagate(False)

        def GoAepF():
            AepF = Frame(AppF, height=530, width=800)
            AepF.grid(row=1, column=0, columnspan=8)
            AepF.grid_propagate(False)

            Anavbarbtns = [Abtn1, Abtn2]
            for x in Anavbarbtns:
                x.configure(bg='#D4BEAD')
            Abtn1.configure(bg='#A07858')

            def outlookinbox():
                try:
                    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
                    inbox = outlook.GetDefaultFolder(6)
                    messages = inbox.Items
                    Aeptxt.configure(state=NORMAL)
                    Aeptxt.insert(END, "\n\nInbox: Microsoft Outlook:\n\n")
                    for x in messages:
                        Aeptxt.insert(END,
                                    "From \t\t" + str(x.sender) + "\t-\t" + str(x.sender.Address) + "\n" +
                                    "To\t\t" + str(x.To) + "\n" +
                                    "Time\t\t" + str(x.ReceivedTime) + "\n" +
                                    "Subject\t\t" + str(x.subject) + "\n" +
                                    #"Body\n" + str(x.body) +
                                    "\n\n")
                    Aeptxt.configure(state=DISABLED)
                except AttributeError:
                    pass
                except Exception as e:
                    messagebox.showinfo('Error', e)
            def Aepexport():
                try:
                    AepReport = open('Reports/Application_Level.txt', 'a', encoding="utf-8")
                    AepReport.write("Time exported : " + str(datetime.now()) +
                                     "\n\n__INBOX - MICROSOFT OUTLOOK__\n\n")
                    response = Aeptxt.get("1.0", END)
                    for x in response:
                        AepReport.write(x)
                    AepReport.close()
                    path = os.path.dirname(os.path.realpath('Reports/Application_Level.txt'))
                    messagebox.showinfo('Export Complete',
                                        'Your records has been successfully exported to: \n' + path + '\\' + 'Application_Level.txt')
                    md5sum("Reports/Application_Level.txt")
                except Exception as e:
                    messagebox.showinfo('Error', e)

            Aepbtn1 = Button(AepF, text='Outlook', height=2, width=20, command=outlookinbox, bg='gray77')
            Aepbtn2 = Button(AepF, text='Export', height=2, width=12, command=Aepexport, bg='gray77')

            Aepscrolly = Scrollbar(AepF, orient=VERTICAL)
            Aepscrollx = Scrollbar(AepF, orient=HORIZONTAL)
            Aeptxt = Text(AepF, yscrollcommand=Aepscrolly.set, xscrollcommand=Aepscrollx.set, height=22, width=90, wrap='none')
            Aepscrolly.configure(comman=Aeptxt.yview)
            Aepscrollx.configure(comman=Aeptxt.xview)
            Aeptxt.configure(state=DISABLED)

            Aepbtn1.grid(row=1, column=0, padx=20, pady=20, sticky='ew')
            Aeptxt.grid(row=3, column=0, columnspan=4, sticky='we', padx=(20, 0))
            Aepscrolly.grid(row=3, column=4, sticky='nsw')
            Aepscrollx.grid(row=4, column=0, columnspan=4, sticky='we', padx=(20, 0))
            Aepbtn2.grid(row=5, column=2, padx=10, pady=20, sticky='ew')

        def GoAwhF():
            AwhF = Frame(AppF, height=530, width=800)
            AwhF.grid(row=1, column=0, columnspan=8)
            AwhF.grid_propagate(False)

            Anavbarbtns = [Abtn1, Abtn2]
            for x in Anavbarbtns:
                x.configure(bg='#D4BEAD')
            Abtn2.configure(bg='#A07858')

            def chromehis():
                try:
                    f = Chrome()
                    outputs = f.fetch_history()
                    his = outputs.histories
                    Awhtxt.configure(state=NORMAL)
                    for x in his:
                        Awhtxt.insert(END, "Time\t:" + str(x[0]) + "\nURL\t:" + str(x[1]) + "\n\n")
                    Awhtxt.configure(state=DISABLED)
                except Exception as e:
                    messagebox.showinfo('Error', e)

            def Awhexport():
                try:
                    AwhReport = open('Reports/Application_Level.txt', 'a', encoding="utf-8")
                    AwhReport.write("Time exported : " + str(datetime.now()) +
                                     "\n\n__CHROME HISTORY__\n\n")
                    response = Awhtxt.get("1.0", END)
                    for x in response:
                        AwhReport.write(x)
                    AwhReport.close()
                    path = os.path.dirname(os.path.realpath('Reports/Application_Level.txt'))
                    messagebox.showinfo('Export Complete',
                                        'Your records has been successfully exported to: \n' + path + '\\' + 'Application_Level.txt')
                    md5sum("Reports/Application_Level.txt")
                except Exception as e:
                    messagebox.showinfo('Error', e)

            Awhbtn1 = Button(AwhF, text='Google Chrome', height=2, width=20, command=chromehis, bg='gray77')
            Awhbtn2 = Button(AwhF, text='Export', height=2, width=12, command=Awhexport, bg='gray77')

            Awhscrolly = Scrollbar(AwhF, orient=VERTICAL)
            Awhscrollx = Scrollbar(AwhF, orient=HORIZONTAL)
            Awhtxt = Text(AwhF, yscrollcommand=Awhscrolly.set, xscrollcommand=Awhscrollx.set, height=22, width=90, wrap='none')
            Awhscrolly.configure(comman=Awhtxt.yview)
            Awhscrollx.configure(comman=Awhtxt.xview)
            Awhtxt.configure(state=DISABLED)

            Awhbtn1.grid(row=1, column=0, padx=20, pady=20, sticky='ew')
            Awhtxt.grid(row=3, column=0, columnspan=4, sticky='we', padx=(20, 0))
            Awhscrolly.grid(row=3, column=4, sticky='nsw')
            Awhscrollx.grid(row=4, column=0, columnspan=4, sticky='we', padx=(20, 0))
            Awhbtn2.grid(row=5, column=2, padx=10, pady=20, sticky='ew')

        Abtn1 = Button(AppF, text='Email Parser', height=3, width=55, command=GoAepF, bg='#D4BEAD')
        Abtn2 = Button(AppF, text='Web Browser History', height=3, width=56, command=GoAwhF, bg='#D4BEAD')

        Abtn1.grid(row=0, column=0, sticky='nsew')
        Abtn2.grid(row=0, column=1, sticky='nsew')

        GoAepF()

    def GoWindows():
        for widget in Tabs.winfo_children():
            widget.destroy()
        navbarbtns = [devicebtn, appbtn, windbtn, multbtn, logbtn, encryptbtn, sendbtn]
        for x in navbarbtns:
            x.configure(bg='gray77')
        windbtn.configure(bg='#885933')
        WindowsF = Frame(Tabs, height=600, width=800)
        WindowsF.grid(row=0, column=0)
        WindowsF.grid_propagate(False)

        def GoWevtF():
            WevtF = Frame(WindowsF, height=500, width=800)
            WevtF.grid(row=2, column=0, columnspan=8)
            WevtF.grid_propagate(False)

            Wnavbarbtns = [Wbtn1, Wbtn2, Wbtn3, Wbtn4, Wbtn5, Wbtn6]
            for x in Wnavbarbtns:
                x.configure(bg='#D4BEAD')
            Wbtn1.configure(bg='#A07858')

            def logOn():
                try:
                    server = 'localhost'
                    logtype = 'Security'
                    hand = win32evtlog.OpenEventLog(server, logtype)
                    flags = win32evtlog.EVENTLOG_SEQUENTIAL_READ | win32evtlog.EVENTLOG_BACKWARDS_READ
                    events = win32evtlog.ReadEventLog(hand, flags, 0)
                    x = 1
                    Wevttxt.configure(state=NORMAL)
                    while x < 10000:
                        x += 1
                        events = win32evtlog.ReadEventLog(hand, flags, 0)
                        for event in events:
                            if event.EventID == 4624:
                                Wevttxt.insert(END, "LogOn\t" + str(event.EventID) + "\t" + str(event.ComputerName) + "\t" + str(
                                        event.TimeGenerated) + "\t" + str(event.SourceName) + "\n")
                    Wevttxt.configure(state=DISABLED)
                except Exception as e:
                    messagebox.showinfo('Error', e)

            def logOff():
                try:
                    server = 'localhost'
                    logtype = 'Security'
                    hand = win32evtlog.OpenEventLog(server, logtype)
                    flags = win32evtlog.EVENTLOG_SEQUENTIAL_READ | win32evtlog.EVENTLOG_BACKWARDS_READ
                    events = win32evtlog.ReadEventLog(hand, flags, 0)
                    x = 1
                    Wevttxt.configure(state=NORMAL)
                    while x < 10000:
                        x += 1
                        events = win32evtlog.ReadEventLog(hand, flags, 0)
                        for event in events:
                            if event.EventID == 4647:
                                Wevttxt.insert(END, "LogOff\t" + str(event.EventID) + "\t" + str(event.ComputerName) + "\t" + str(
                                        event.TimeGenerated) + "\t" + str(event.SourceName) + "\n")
                    Wevttxt.configure(state=DISABLED)
                except Exception as e:
                    messagebox.showinfo('Error', e)

            def failedLog():
                try:
                    server = 'localhost'
                    logtype = 'Security'
                    hand = win32evtlog.OpenEventLog(server, logtype)
                    flags = win32evtlog.EVENTLOG_SEQUENTIAL_READ | win32evtlog.EVENTLOG_BACKWARDS_READ
                    events = win32evtlog.ReadEventLog(hand, flags, 0)
                    x = 1
                    Wevttxt.configure(state=NORMAL)
                    while x < 10000:
                        x += 1
                        events = win32evtlog.ReadEventLog(hand, flags, 0)
                        for event in events:
                            if event.EventID == 4625:
                                Wevttxt.insert(END, "Failed Login\t" + str(event.EventID) + "\t" + str(event.ComputerName) + "\t" + str(
                                        event.TimeGenerated) + "\t" + str(event.SourceName) + "\n")
                    Wevttxt.configure(state=DISABLED)
                except Exception as e:
                    messagebox.showinfo('Error', e)

            def Wevtexport():
                try:
                    WevtReport = open('Reports/Windows_Artefacts.txt', 'a', encoding="utf-8")
                    WevtReport.write("Time exported : " + str(datetime.now()) +
                                     "\n\n__WINDOWS EVENT LOGS__\n\n")
                    response = Wevttxt.get("1.0", END)
                    for x in response:
                        WevtReport.write(x)
                    WevtReport.close()
                    path = os.path.dirname(os.path.realpath('Reports/Windows_Artefacts.txt'))
                    messagebox.showinfo('Export Complete',
                                        'Your records has been successfully exported to: \n' + path + '\\' + 'Windows_Artefacts.txt')
                    md5sum("Reports/Windows_Artefacts.txt")
                except Exception as e:
                    messagebox.showinfo('Error', e)

            Wevtbtn1 = Button(WevtF, text='LogOn', height=2, width=20, command=logOn, bg='gray77')
            Wevtbtn2 = Button(WevtF, text='LogOff', height=2, width=12, command=logOff, bg='gray77')
            Wevtbtn3 = Button(WevtF, text='Failed Logins', height=2, width=20, command=failedLog, bg='gray77')
            Wevtbtn4 = Button(WevtF, text='Export', height=2, width=12, command=Wevtexport, bg='gray77')

            Wevtscrolly = Scrollbar(WevtF, orient=VERTICAL)
            Wevtscrollx = Scrollbar(WevtF, orient=HORIZONTAL)
            Wevttxt = Text(WevtF, yscrollcommand=Wevtscrolly.set, xscrollcommand=Wevtscrollx.set, height=21, width=95, wrap='none')
            Wevtscrolly.configure(command=Wevttxt.yview)
            Wevtscrollx.configure(command=Wevttxt.xview)
            Wevttxt.configure(state=DISABLED)

            Wevtbtn1.grid(row=2, column=0, padx=20, pady=10, sticky='ew')
            Wevtbtn2.grid(row=2, column=1, padx=20, pady=10, sticky='ew')
            Wevtbtn3.grid(row=2, column=2, padx=(20, 0), pady=10, sticky='ew')
            Wevttxt.grid(row=3, column=0, columnspan=3, padx=(20, 0), sticky='e')
            Wevtscrolly.grid(row=3, column=3, sticky='nsw')
            Wevtscrollx.grid(row=4, column=0, columnspan=3, padx=(20, 0), sticky='ew')
            Wevtbtn4.grid(row=5, column=2, padx=10, pady=10, sticky='ew')

        def GoWrbF():
            WrbF = Frame(WindowsF, height=500, width=800)
            WrbF.grid(row=2, column=0, columnspan=8)
            WrbF.grid_propagate(False)

            Wnavbarbtns = [Wbtn1, Wbtn2, Wbtn3, Wbtn4, Wbtn5, Wbtn6]
            for x in Wnavbarbtns:
                x.configure(bg='#D4BEAD')
            Wbtn2.configure(bg='#A07858')

            def recyclebin():
                try:
                    r = list(winshell.recycle_bin())
                    Wrbtxt.configure(state=NORMAL)
                    for index, value in enumerate(r):
                        txt = (value.recycle_date(), "\t", value.original_filename()), "\n"
                        Wrbtxt.insert(1.0, txt)
                    Wrbtxt.configure(state=DISABLED)
                except Exception as e:
                    messagebox.showinfo('Error', e)

            def Wrbexport():
                try:
                    WrbReport = open('Reports/Windows_Artefacts.txt', 'a', encoding="utf-8")
                    WrbReport.write("Time exported : " + str(datetime.now()) +
                                     "\n\n__DELETED ITEMS__\n\n")
                    response = Wrbtxt.get("1.0", END)
                    for x in response:
                        WrbReport.write(x)
                    WrbReport.close()
                    path = os.path.dirname(os.path.realpath('Reports/Windows_Artefacts.txt'))
                    messagebox.showinfo('Export Complete',
                                        'Your records has been successfully exported to: \n' + path + '\\' + 'Windows_Artefacts.txt')
                    md5sum("Reports/Windows_Artefacts.txt")
                except Exception as e:
                    messagebox.showinfo('Error', e)

            Wrbbtn1 = Button(WrbF, text='List Deleted Items', height=2, width=20,  command=recyclebin, bg='gray77')
            Wrbbtn2 = Button(WrbF, text='Export', height=2, width=12, command=Wrbexport, bg='gray77')

            Wrbscrolly = Scrollbar(WrbF, orient=VERTICAL)
            Wrbscrollx = Scrollbar(WrbF, orient=HORIZONTAL)
            Wrbtxt = Text(WrbF, yscrollcommand=Wrbscrolly.set, xscrollcommand=Wrbscrollx.set, height=21, wrap='none')
            Wrbscrolly.configure(command=Wrbtxt.yview)
            Wrbscrollx.configure(command=Wrbtxt.xview)
            Wrbtxt.configure(state=DISABLED)

            Wrbbtn1.grid(row=2, column=0, padx=50, pady=10, sticky='ew')
            Wrbtxt.grid(row=3, column=0, columnspan=3, padx=(50, 0), sticky='e')
            Wrbscrolly.grid(row=3, column=3, sticky='nsw')
            Wrbscrollx.grid(row=4, column=0, columnspan=3, padx=(50, 0), sticky='we')
            Wrbbtn2.grid(row=5, column=2, padx=10, pady=10, sticky='ew')

        def GoWtnF():
            WtnF = Frame(WindowsF, height=500, width=800)
            WtnF.grid(row=2, column=0, columnspan=8)
            WtnF.grid_propagate(False)

            Wnavbarbtns = [Wbtn1, Wbtn2, Wbtn3, Wbtn4, Wbtn5, Wbtn6]
            for x in Wnavbarbtns:
                x.configure(bg='#D4BEAD')
            Wbtn3.configure(bg='#A07858')

            def examinethumb():
                try:
                    input_file = Wtntxt.get()
                    exe = "D:/Final Project/pythonProject/thumbcache_viewer.exe"
                    process = subprocess.Popen([exe, input_file])
                except Exception as e:
                    messagebox.showinfo('Error', e)

            def addfile():
                try:
                    user = str(os.getlogin())
                    addfile.filename = filedialog.askopenfilename(initialdir="C:/Users/%s/AppData/Local/Microsoft/Windows/Explorer/" % user, title="Select a File", filetypes=(("Databases", "*.db"),("all files","*,*")))
                    Wtntxt.configure(state=NORMAL)
                    Wtntxt.delete(0, END)
                    Wtntxt.insert(0, addfile.filename)
                    Wtntxt.configure(state=DISABLED)
                except Exception as e:
                    messagebox.showinfo('Error', e)

            Wtnlbl = Label(WtnF, text='Add File')
            Wtntxt = Entry(WtnF, width=100, font=Font2)
            Wtnbtn1 = Button(WtnF, text='Browse', height=2, width=12, command=addfile, bg='gray77')
            Wtnbtn2 = Button(WtnF, text='Examine', height=2, width=12, command=examinethumb, bg='gray77')
            Wtntxt.configure(state=DISABLED)
            Wtnlbl2 = Label(WtnF)

            Wtnlbl.grid(row=0, column=0, padx=(10, 10), pady=(200, 10))
            Wtntxt.grid(row=0, column=1, columnspan=2, padx=10, pady=(200, 10))
            Wtnbtn1.grid(row=1, column=1, padx=10, pady=10, sticky='ew')
            Wtnbtn2.grid(row=1, column=2, padx=10, pady=10, sticky='ew')
            Wtnlbl2.grid(row=3, column=0, columnspan=10)

        def GoWraF():
            WraF = Frame(WindowsF, height=500, width=800)
            WraF.grid(row=2, column=0, columnspan=8)
            WraF.grid_propagate(False)

            Wnavbarbtns = [Wbtn1, Wbtn2, Wbtn3, Wbtn4, Wbtn5, Wbtn6]
            for x in Wnavbarbtns:
                x.configure(bg='#D4BEAD')
            Wbtn4.configure(bg='#A07858')

            def recentfiles():
                try:
                    user = str(os.getlogin())
                    arr = os.listdir("C:/Users/%s/AppData/Roaming/Microsoft/Windows/Recent" % user)
                    Wratxt.configure(state=NORMAL)
                    for x in arr:
                        Wratxt.insert(END, x + "\n")
                    Wratxt.configure(state=DISABLED)
                except Exception as e:
                    messagebox.showinfo('Error', e)

            def browserecent():
                try:
                    user = str(os.getlogin())
                    os.startfile("C:/Users/%s/AppData/Roaming/Microsoft/Windows/Recent" % user)
                except Exception as e:
                    messagebox.showinfo('Error', e)

            def Wraexport():
                try:
                    WraReport = open('Reports/Windows_Artefacts.txt', 'a', encoding="utf-8")
                    WraReport.write("Time exported : " + str(datetime.now()) +
                                     "\n\n__RECENT ACTIVITIES__\n\n")
                    response = Wratxt.get("1.0", END)
                    for x in response:
                        WraReport.write(x)
                    WraReport.close()
                    path = os.path.dirname(os.path.realpath('Reports/Windows_Artefacts.txt'))
                    messagebox.showinfo('Export Complete',
                                        'Your records has been successfully exported to: \n' + path + '\\' + 'Windows_Artefacts.txt')
                    md5sum("Reports/Windows_Artefacts.txt")
                except Exception as e:
                    messagebox.showinfo('Error', e)

            Wrabtn1 = Button(WraF, text='List Recent Activities', height=2, width=20, command=recentfiles, bg='gray77')
            Wrabtn2 = Button(WraF, text='Open in File Explorer', height=2, width=12, command=browserecent, bg='gray77')
            Wrabtn3 = Button(WraF, text='Export', height=2, width=12, command=Wraexport, bg='gray77')

            Wrascrolly = Scrollbar(WraF, orient=VERTICAL)
            Wrascrollx = Scrollbar(WraF, orient=HORIZONTAL)
            Wratxt = Text(WraF, yscrollcommand=Wrascrolly.set, xscrollcommand=Wrascrollx.set, height=21, wrap='none')
            Wrascrolly.configure(command=Wratxt.yview)
            Wrascrollx.configure(command=Wratxt.xview)
            Wratxt.configure(state=DISABLED)

            Wrabtn1.grid(row=2, column=0, padx=50, pady=10, sticky='ew')
            Wratxt.grid(row=3, column=0, columnspan=3, padx=(50, 0), sticky='e')
            Wrascrolly.grid(row=3, column=3, sticky='nsw')
            Wrascrollx.grid(row=4, column=0, columnspan=3, padx=(50, 0), sticky='we')
            Wrabtn2.grid(row=5, column=1, padx=10, pady=10, sticky='ew')
            Wrabtn3.grid(row=5, column=2, padx=10, pady=10, sticky='ew')

        def GoWpreF():
            WpreF = Frame(WindowsF, height=500, width=800)
            WpreF.grid(row=2, column=0, columnspan=8)
            WpreF.grid_propagate(False)

            Wnavbarbtns = [Wbtn1, Wbtn2, Wbtn3, Wbtn4, Wbtn5, Wbtn6]
            for x in Wnavbarbtns:
                x.configure(bg='#D4BEAD')
            Wbtn5.configure(bg='#A07858')

            def prefetch():
                try:
                    prefetch_directory = "C:\Windows\Prefetch\\"
                    prefetch_files = os.listdir(prefetch_directory)
                    Wpretxt.configure(state=NORMAL)
                    for pf_file in prefetch_files:
                        if pf_file[-2:] == "pf":
                            full_path = prefetch_directory + pf_file

                            app_name = pf_file[:-12]

                            first_executed = os.path.getctime(full_path)
                            first_executed = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(first_executed))

                            last_executed = os.path.getmtime(full_path)
                            last_executed = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(last_executed))

                            result = "Programe name : " + app_name + "\n\tFirst Executed\t:" + first_executed + "\n\tLast Executed\t:" + last_executed + "\n\tPrefetch\t:" + pf_file + "\n\n"
                            Wpretxt.insert(END, result)
                    Wpretxt.configure(state=DISABLED)
                except Exception as e:
                    messagebox.showinfo('Error', e)

            def Wpreexport():
                try:
                    WpreReport = open('Reports/Windows_Artefacts.txt', 'a', encoding="utf-8")
                    WpreReport.write("Time exported : " + str(datetime.now()) +
                                     "\n\n__PREFETCH FILES__\n\n")
                    response = Wpretxt.get("1.0", END)
                    for x in response:
                        WpreReport.write(x)
                    WpreReport.close()
                    path = os.path.dirname(os.path.realpath('Reports/Windows_Artefacts.txt'))
                    messagebox.showinfo('Export Complete',
                                        'Your records has been successfully exported to: \n' + path + '\\' + 'Windows_Artefacts.txt')
                    md5sum("Reports/Windows_Artefacts.txt")
                except Exception as e:
                    messagebox.showinfo('Error', e)

            Wprebtn1 = Button(WpreF, text='Examine', height=2, width=20, command=prefetch, bg='gray77')
            Wprebtn2 = Button(WpreF, text='Export', height=2, width=12, command=Wpreexport, bg='gray77')

            Wprescrolly = Scrollbar(WpreF, orient=VERTICAL)
            Wprescrollx = Scrollbar(WpreF, orient=HORIZONTAL)
            Wpretxt = Text(WpreF, yscrollcommand=Wprescrolly.set, xscrollcommand=Wprescrollx.set, height=21, wrap='none')
            Wprescrolly.configure(command=Wpretxt.yview)
            Wprescrollx.configure(command=Wpretxt.xview)
            Wpretxt.configure(state=DISABLED)

            Wprebtn1.grid(row=2, column=0, padx=50, pady=10, sticky='ew')
            Wpretxt.grid(row=3, column=0, columnspan=3, padx=(50, 0), sticky='e')
            Wprescrolly.grid(row=3, column=3, sticky='nsw')
            Wprescrollx.grid(row=4, column=0, columnspan=3, padx=(50, 0), sticky='we')
            Wprebtn2.grid(row=5, column=2, padx=10, pady=10, sticky='ew')

        def GoWappF():
            WappF = Frame(WindowsF, height=500, width=800)
            WappF.grid(row=2, column=0, columnspan=8)
            WappF.grid_propagate(False)

            Wnavbarbtns = [Wbtn1, Wbtn2, Wbtn3, Wbtn4, Wbtn5, Wbtn6]
            for x in Wnavbarbtns:
                x.configure(bg='#D4BEAD')
            Wbtn6.configure(bg='#A07858')

            def applist():
                try:
                    path = "reg query \"HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\" /s | findstr /B \".*DisplayName\""
                    path2 = "reg query \"HKLM\\Software\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall\" /s | findstr /B \".*DisplayName\""
                    response = os.popen(path)
                    response2 = os.popen(path2)
                    Wapptxt.configure(state=NORMAL)
                    for x in response:
                        Wapptxt.insert(END, x)
                    for x in response2:
                        Wapptxt.insert(END, x)
                    Wapptxt.configure(state=DISABLED)
                except Exception as e:
                    messagebox.showinfo('Error', e)

            def msapplist():
                try:
                    exe = 'powershell'
                    path = "Get-AppxPackage –AllUsers | Select Name, PackageFullName"
                    response = subprocess.Popen([exe, path], stdout=subprocess.PIPE)
                    Wapptxt.configure(state=NORMAL)
                    for x in response.stdout:
                        Wapptxt.insert(END, x)
                    Wapptxt.configure(state=DISABLED)
                except Exception as e:
                    messagebox.showinfo('Error', e)

            def Wappexport():
                try:
                    WappReport = open('Reports/Windows_Artefacts.txt', 'a', encoding="utf-8")
                    WappReport.write("Time exported : " + str(datetime.now()) +
                                     "\n\n__APP LIST__\n\n")
                    response = Wapptxt.get("1.0", END)
                    for x in response:
                        WappReport.write(x)
                    WappReport.close()
                    path = os.path.dirname(os.path.realpath('Reports/Windows_Artefacts.txt'))
                    messagebox.showinfo('Export Complete',
                                        'Your records has been successfully exported to: \n' + path + '\\' + 'Windows_Artefacts.txt')
                    md5sum("Reports/Windows_Artefacts.txt")
                except Exception as e:
                    messagebox.showinfo('Error', e)

            Wappbtn1 = Button(WappF, text='All Apps Installed', height=2, width=20, command=applist, bg='gray77')
            Wappbtn2 = Button(WappF, text='Microsoft Apps Installed', height=2, width=12, command=msapplist, bg='gray77')
            Wappbtn3 = Button(WappF, text='Export', height=2, width=12, command=Wappexport, bg='gray77')

            Wappscrolly = Scrollbar(WappF, orient=VERTICAL)
            Wappscrollx = Scrollbar(WappF, orient=HORIZONTAL)
            Wapptxt = Text(WappF, yscrollcommand=Wappscrolly.set, xscrollcommand=Wappscrollx.set, height=21, width=90, wrap='none')
            Wappscrolly.configure(command=Wapptxt.yview)
            Wappscrollx.configure(command=Wapptxt.xview)
            Wapptxt.configure(state=DISABLED)

            Wappbtn1.grid(row=2, column=0, padx=20, pady=10, sticky='ew')
            Wappbtn2.grid(row=2, column=1, padx=10, pady=10, sticky='ew')
            Wapptxt.grid(row=3, column=0, columnspan=2, padx=(20, 0), sticky='e')
            Wappscrolly.grid(row=3, column=2, sticky='nsw')
            Wappscrollx.grid(row=4, column=0, columnspan=2, padx=(20, 0), sticky='we')
            Wappbtn3.grid(row=5, column=1, padx=10, pady=10, sticky='ew')

        Wbtn1 = Button(WindowsF, text='Event Log Parser', height=3, width=37, command=GoWevtF, bg='#D4BEAD')
        Wbtn2 = Button(WindowsF, text='Recycle Bin ', height=3, width=37, command=GoWrbF, bg='#D4BEAD')
        Wbtn3 = Button(WindowsF, text='Thumbnails', height=3, width=36, command=GoWtnF, bg='#D4BEAD')
        Wbtn4 = Button(WindowsF, text='Recent Activities', height=3, width=37, command=GoWraF, bg='#D4BEAD')
        Wbtn5 = Button(WindowsF, text='Prefetch Parser', height=3, width=37, command=GoWpreF, bg='#D4BEAD')
        Wbtn6 = Button(WindowsF, text='App List', height=3, width=36, command=GoWappF, bg='#D4BEAD')

        Wbtn1.grid(row=0, column=0, sticky='ew')
        Wbtn2.grid(row=0, column=1, sticky='ew')
        Wbtn3.grid(row=0, column=2, sticky='ew')
        Wbtn4.grid(row=1, column=0, sticky='ew')
        Wbtn5.grid(row=1, column=1, sticky='ew')
        Wbtn6.grid(row=1, column=2, sticky='ew')

        GoWevtF()

    def GoMeta():
        for widget in Tabs.winfo_children():
            widget.destroy()
        navbarbtns = [devicebtn, appbtn, windbtn, multbtn, logbtn, encryptbtn, sendbtn]
        for x in navbarbtns:
            x.configure(bg='gray77')
        multbtn.configure(bg='#885933')

        MetaF = Frame(Tabs, height=600, width=800)
        MetaF.grid(row=0, column=0)
        MetaF.grid_propagate(False)

        def examinemedia():
            try:
                input_file = Mtxt1.get()
                exe = "D:/Softwares/exiftool-12.34/exiftool.exe"
                process = subprocess.Popen([exe, input_file], stdout=subprocess.PIPE, stderr=subprocess.STDOUT, universal_newlines=True)
                Mtxt2.configure(state=NORMAL)
                Mtxt2.insert(END, "File: " + input_file + "\n\n")
                for output in process.stdout:
                    Mtxt2.insert(END, output.strip() + "\n")
                Mtxt2.insert(END, "\n\n")
                Mtxt2.configure(state=DISABLED)
            except Exception as e:
                messagebox.showinfo('Error', e)

        def addfile():
            try:
                addfile.filename = filedialog.askopenfilename(title="Select a File")
                Mtxt1.configure(state=NORMAL)
                Mtxt1.delete(0, END)
                Mtxt1.insert(0, addfile.filename)
                Mtxt1.configure(state=DISABLED)
            except Exception as e:
                messagebox.showinfo('Error', e)

        def Mexport():
            try:
                MReport = open('Reports/Metadata_Analysis.txt', 'a', encoding="utf-8")
                MReport.write("Time exported : " + str(datetime.now()) +
                                 "\n\n__METADATA ANALYSIS__\n\n")
                response = Mtxt2.get("1.0", END)
                for x in response:
                    MReport.write(x)
                MReport.close()
                path = os.path.dirname(os.path.realpath('Reports/Metadata_Analysis.txt'))
                messagebox.showinfo('Export Complete',
                                    'Your records has been successfully exported to: \n' + path + '\\' + 'Metadata_Analysis.txt')
                md5sum("Reports/Metadata_Analysis.txt")
            except Exception as e:
                messagebox.showinfo('Error', e)

        Mlbl = Label(MetaF, text='Add File')
        Mtxt1 = Entry(MetaF, width=80, font=Font2)
        Mtxt1.configure(state=DISABLED)
        Mbtn1 = Button(MetaF, text='Browse', height=2, width=12, command=addfile, bg='gray77')
        Mbtn2 = Button(MetaF, text='Examine', height=2, width=12, command=examinemedia, bg='gray77')
        Mbtn3 = Button(MetaF, text='Export', height=2, width=12, command=Mexport, bg='gray77')

        Mscrolly = Scrollbar(MetaF, orient=VERTICAL)
        Mscrollx = Scrollbar(MetaF, orient=HORIZONTAL)
        Mtxt2 = Text(MetaF, yscrollcommand=Mscrolly.set, xscrollcommand=Mscrollx.set, height=21, wrap='none')
        Mscrolly.configure(command=Mtxt2.yview)
        Mscrollx.configure(command=Mtxt2.xview)
        Mtxt2.configure(state=DISABLED)

        Mlbl.grid(row=0, column=0, padx=(10, 10), pady=(50, 10))
        Mtxt1.grid(row=0, column=1, columnspan=2, pady=(50, 10))
        Mbtn1.grid(row=1, column=1, sticky='ew', padx=20, pady=(0, 50))
        Mbtn2.grid(row=1, column=2, sticky='ew', padx=20, pady=(0, 50))
        Mtxt2.grid(row=3, column=0, columnspan=3, sticky='e', padx=(50, 0))
        Mscrolly.grid(row=3, column=3, sticky='nsw')
        Mscrollx.grid(row=4, column=0, columnspan=3, sticky='we', padx=(50, 0))
        Mbtn3.grid(row=5, column=2, sticky='ew', padx=10, pady=10)

    def GoLog():
        for widget in Tabs.winfo_children():
            widget.destroy()
        navbarbtns = [devicebtn, appbtn, windbtn, multbtn, logbtn, encryptbtn, sendbtn]
        for x in navbarbtns:
            x.configure(bg='gray77')
        logbtn.configure(bg='#885933')

        LogF = Frame(Tabs, height=600, width=800)
        LogF.grid(row=0, column=0)
        LogF.grid_propagate(False)

        Lscrolly = Scrollbar(LogF, orient=VERTICAL)
        Lscrollx = Scrollbar(LogF, orient=HORIZONTAL)
        Ltxt = Text(LogF, yscrollcommand=Lscrolly.set, xscrollcommand=Lscrollx.set, height=34, width=93, wrap='none')
        Lscrolly.configure(command=Ltxt.yview)
        Lscrollx.configure(command=Ltxt.xview)
        Ltxt.insert(1.0, "Log File\n ")
        Ltxt.configure(state=DISABLED)

        def ShowLogData():
            try:
                Ltxt.configure(state=NORMAL)
                Ltxt.delete(1.0, END)
                Ltxt.insert(1.0, "\tReport Name\t\t\tExport Date & Time\t\t\t\tMD5\t\t\tPIN\n\n")
                cursor.execute("SELECT * FROM reporthash")
                rows = cursor.fetchall()
                for row in rows:
                    Ltxt.insert(END, row)
                    Ltxt.insert(END, "\n\n")
                Ltxt.configure(state=DISABLED)
            except Exception as e:
                messagebox.showinfo('Error', e)

        Ltxt.grid(row=0, column=0, columnspan=3, sticky='e', padx=(10, 0), pady=(20, 0))
        Lscrolly.grid(row=0, column=3, sticky='nsw', pady=(20, 0))
        Lscrollx.grid(row=1, column=0, columnspan=3, sticky='we', padx=(10, 0))
        ShowLogData()

    def GoEncrypt():
        for widget in Tabs.winfo_children():
            widget.destroy()
        navbarbtns = [devicebtn, appbtn, windbtn, multbtn, logbtn, encryptbtn, sendbtn]
        for x in navbarbtns:
            x.configure(bg='gray77')
        encryptbtn.configure(bg='#885933')

        EncryptF = Frame(Tabs, height=600, width=800)
        EncryptF.grid(row=0, column=0)
        EncryptF.grid_propagate(False)

        global filename

        def Eaddfile():
            try:
                Eaddfile.filename = filedialog.askopenfilename(initialdir="Reports",
                                                                 title="Select a File",
                                                                 filetypes=(("Text Files", "*.txt"), ("all files", "*,*")))
                Etxt1.configure(state=NORMAL)
                Etxt1.delete(0, END)
                Etxt1.insert(0, Eaddfile.filename)
                Etxt1.configure(state=DISABLED)
            except Exception as e:
                messagebox.showinfo('Error', e)

        def encrypt():
            tmp = Etxt2.get()
            try:
                int(tmp)
                if len(Etxt1.get()) == 0:
                    Elbl3.configure(text="Please Browse a File To Encrypt")
                elif not len(tmp) == 3:
                    Elbl3.configure(text="Please Enter a 3 number PIN")
                elif type(tmp) == int:
                    Elbl3.configure(text="Please Enter a 3 number PIN")
                else:
                    tmp = ''.join(e for e in tmp if e.isalnum())
                    key = tmp + ("s" * (43 - len(tmp)) + "=")

                    fernet = Fernet(key)

                    with open(Eaddfile.filename, 'rb') as file:
                        original = file.read()
                        encrypted = fernet.encrypt(original)

                    with open(Eaddfile.filename, 'wb') as encrypted_file:
                        encrypted_file.write(encrypted)
                    messagebox.showinfo('', 'Encryption Complete')

                    pin = str(Etxt2.get())
                    fname = os.path.basename(Eaddfile.filename)
                    pininlog(fname, pin)
            except ValueError:
                Elbl3.configure(text="Please Enter  a 3 number PIN")
            Etxt2.delete(0, END)
            Etxt1.configure(state=NORMAL)
            Etxt1.delete(0, END)
            Etxt1.configure(state=DISABLED)

        def Daddfile():
            try:
                Daddfile.filename = filedialog.askopenfilename(initialdir="Reports",
                                                               title="Select a File",
                                                               filetypes=(("Text Files", "*.txt"), ("all files", "*,*")))
                Dtxt1.configure(state=NORMAL)
                Dtxt1.delete(0, END)
                Dtxt1.insert(0, Daddfile.filename)
                Dtxt1.configure(state=DISABLED)
            except Exception as e:
                messagebox.showinfo('Error', e)

        def getpin():
            try:
                conn = sqlite3.connect("k9DB.db")
                fname = os.path.basename(Daddfile.filename)
                cursor = conn.cursor()
                cursor.execute("SELECT pin FROM reporthash WHERE name = ?", (fname,))
                rows = cursor.fetchone()
                for row in rows:
                    return (row)
            except Exception as e:
                messagebox.showinfo('Error', e)

        def decrypt():
            try:
                Dlbl3.configure(text="")
                tmp = str(Dtxt2.get())
                if len(Dtxt1.get()) == 0:
                    Dlbl3.configure(text="Please Browse a File To Encrypt")
                elif tmp == getpin():
                    tmp = ''.join(e for e in tmp if e.isalnum())
                    key = tmp + ("s" * (43 - len(tmp)) + "=")

                    fernet = Fernet(key)

                    with open(Daddfile.filename, 'rb') as enc_file:
                        encrypted = enc_file.read()
                        decrypted = fernet.decrypt(encrypted)

                    with open(Daddfile.filename, 'wb') as dec_file:
                        dec_file.write(decrypted)
                        messagebox.showinfo('', 'Decryption Complete')

                    fname = os.path.basename(Daddfile.filename)
                    removepin(fname)
                else:
                    Dlbl3.configure(text="Wrong Pin")
                Dtxt2.delete(0, END)
                Dtxt1.configure(state=NORMAL)
                Dtxt1.delete(0, END)
                Dtxt1.configure(state=DISABLED)
            except Exception as e:
                messagebox.showinfo('Error', e)

        Elbl1 = Label(EncryptF, text='Add File')
        Etxt1 = Entry(EncryptF, width=80, font=Font2)
        Ebtn1 = Button(EncryptF, text='Browse', height=2, width=12, command=Eaddfile, bg='gray77')
        Elbl2 = Label(EncryptF, text='PIN')
        Etxt2 = Entry(EncryptF, width=40, font=Font2)
        Ebtn2 = Button(EncryptF, text='Encrypt', height=2, width=12, command=encrypt, bg='gray77')
        Etxt1.configure(state=DISABLED)
        Elbl3 = Label(EncryptF, text='Please Enter a 3 number PIN', font=Font2)

        Dlbl1 = Label(EncryptF, text='Add File')
        Dtxt1 = Entry(EncryptF, width=80, font=Font2)
        Dbtn1 = Button(EncryptF, text='Browse', height=2, width=12, command=Daddfile, bg='gray77')
        Dlbl2 = Label(EncryptF, text='PIN')
        Dtxt2 = Entry(EncryptF, width=40, font=Font2)
        Dbtn2 = Button(EncryptF, text='Dencrypt', height=2, width=12, command=decrypt, bg='gray77')
        Dtxt1.configure(state=DISABLED)
        Dlbl3 = Label(EncryptF)

        Elbl1.grid(row=0, column=0, padx=(100, 10), pady=(50, 10))
        Etxt1.grid(row=0, column=1, columnspan=2, padx=10, pady=(50, 10), sticky='w')
        Ebtn1.grid(row=1, column=2, sticky='ew', padx=10, pady=10)
        Elbl2.grid(row=2, column=0, padx=(100, 10), pady=10)
        Etxt2.grid(row=2, column=1, padx=10, pady=10, sticky='w')
        Ebtn2.grid(row=2, column=2, sticky='ew', padx=10, pady=10)
        Elbl3.grid(row=3, column=1)

        Dlbl1.grid(row=4, column=0, padx=(100, 10), pady=(120, 10))
        Dtxt1.grid(row=4, column=1, columnspan=2, padx=10, pady=(120, 10), sticky='w')
        Dbtn1.grid(row=5, column=2, sticky='ew', padx=10, pady=10)
        Dlbl2.grid(row=6, column=0, padx=(100, 10), pady=10)
        Dtxt2.grid(row=6, column=1, padx=10, pady=10, sticky='w')
        Dbtn2.grid(row=6, column=2, sticky='ew', padx=10, pady=10)
        Dlbl3.grid(row=7, column=1)

    def GoSend():
        for widget in Tabs.winfo_children():
            widget.destroy()
        navbarbtns = [devicebtn, appbtn, windbtn, multbtn, logbtn, encryptbtn, sendbtn]
        for x in navbarbtns:
            x.configure(bg='gray77')
        sendbtn.configure(bg='#885933')
        SendF = Frame(Tabs, height=600, width=800)
        SendF.grid(row=0, column=0)
        SendF.grid_propagate(False)

        def send():
            try:
                foldername = "Reports"
                target_dir = "Reports"
                zipobj = zipfile.ZipFile(foldername + '.zip', 'w', zipfile.ZIP_DEFLATED)
                rootlen = len(target_dir) + 1
                for base, dirs, files in os.walk(target_dir):
                    for file in files:
                        fn = os.path.join(base, file)
                        zipobj.write(fn, fn[rootlen:])
                s = smtplib.SMTP('smtp.gmail.com', 587)
                s.starttls()
                s.login("k9.official.app@gmail.com", "sxpfqvhnkjrqvpbm")
                cursor.execute("SELECT * from caseinfo")
                emailid = str(cursor.fetchall()[0][1])
                subject = "Your Files"
                email_content = MIMEMultipart()
                email_content['Subject'] = subject
                attachment = MIMEBase('application', "octet-stream")
                attachment.set_payload(open("Reports.zip", "rb").read())
                encoders.encode_base64(attachment)
                attachment.add_header('Content-Disposition', 'attachment; filename="Reports.zip"')
                email_content.attach(attachment)
                s.sendmail('&&&&&&&&&&&', emailid, email_content.as_string())
            except Exception as e:
                messagebox.showinfo('Error', e)

        Slbl = Label(SendF, text='Email')
        Stxt = Entry(SendF, width=80, font=Font2)
        try:
            cursor.execute("SELECT email FROM caseinfo WHERE id = 1")
            x = cursor.fetchall()
            for row in x:
                Stxt.insert(END, row)
            Stxt.configure(state=DISABLED)
        except Exception as e:
            messagebox.showinfo('Error', e)
        Sbtn = Button(SendF, text='Send Reports', height=2, width=20, command=send, bg='gray77')
        Slbl.grid(row=0, column=0, padx=(100, 10), pady=(200, 10))
        Stxt.grid(row=0, column=1, pady=(200, 10))
        Sbtn.grid(row=2, column=0, columnspan=2, padx=(100, 20))

    homebtn = Button(NavBar, text='Home', height=4, width=27, command=GoHome)
    devicebtn = Button(NavBar, text='Device Analysis', height=5, width=27, command=GoDevice, bg='gray77')
    appbtn = Button(NavBar, text='Application Level Analysis', height=5, width=27, command=GoApp, bg='gray77')
    windbtn = Button(NavBar, text='Windows Artefacts', height=5, width=27, command=GoWindows, bg='gray77')
    multbtn = Button(NavBar, text='Metadata Analysis', height=5, width=27, command=GoMeta, bg='gray77')
    logbtn = Button(NavBar, text='Log File', height=5, width=27, command=GoLog, bg='gray77')
    encryptbtn = Button(NavBar, text='Encrypt Tool', height=5, width=27, command=GoEncrypt, bg='gray77')
    sendbtn = Button(NavBar, text='Send Reports', height=5, width=27, command=GoSend, bg='gray77')

    #homebtn.grid(row=0)
    devicebtn.grid(row=1)
    appbtn.grid(row=2)
    windbtn.grid(row=3)
    multbtn.grid(row=4)
    logbtn.grid(row=5)
    encryptbtn.grid(row=6)
    sendbtn.grid(row=7)

try:
    cursor.execute("SELECT * FROM caseinfo")
    if cursor.fetchall():
        loginScreen()
    else:
        NewCaseScreen()
except Exception as e:
    messagebox.showinfo('Error', e)

K9Root.mainloop()
