import html5lib
import win32com.client
import pythoncom
import re
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from time import sleep
from selenium.common.exceptions import TimeoutException
import dateutil
import datetime
import numpy as np
import logging
import pyodbc
import dateparser

class Handler_Class(object):
    def OnNewMailEx(self, receivedItemsIDs):
        # RecrivedItemIDs is a collection of mail IDs separated by a ",".
        # You know, sometimes more than 1 mail is received at the same moment.

        for ID in receivedItemsIDs.split(","):
            mail = outlook.Session.GetItemFromID(ID)
            subject = mail.Subject
            if subject == "Claims - Claims Form (FNOL)":
                cnxn = pyodbc.connect('DRIVER={SQL Server};server=183.82.0.186;PORT=1433;database=Saxon_Demo;uid=MANOMAY1;pwd=manomay1')
                print("yes")
                msg = mail.Body
                receviedtime = mail.ReceivedTime
                print(receviedtime)
                temp = str(receviedtime).rstrip("+00:00").strip()
                print(temp)
                y= datetime.datetime.strptime(temp, "%Y-%m-%d %H:%M:%S")
                print(y)
                p =y.strftime('%Y-%m-%d %H:%M:%S')
                print(y)
                print(p)
                msghtml = mail.HTMLBody
                dfs = pd.read_html(msghtml,flavor='html5lib')
                df1=dfs[2]
                df = df1.replace(np.nan, '', regex=True)
                print(df)
                df.to_csv("com.csv")
                date_time = df.iloc[17,1]
                #datest = date_time.split()
                #datef = dateutil.parser.parse(date_time)
                t = dateparser.parse(date_time)
                #oldformat = datef
                #datetimeobject = datetime.datetime.strptime(str(oldformat),'%Y-%m-%d  %H:%M:%S')
                newformat = t.strftime('%d/%m/%Y')
                print (newformat)
                newtime = t.strftime('%I:%M %p')
                print(type(newtime))
                print(newtime)
                policy_number = df.iloc[3,1]
                claimname = df.iloc[4,1]
                ad =df.iloc[18,1]
                t = ad.split(r"\D")
                a = re.split('(\d+)',t[0])
                street = a[0] +" "+a[1]
                zipco=a[-3]+a[-2]
                inti = zipco.split(",")[0]
                vechmaker = df.iloc[13,1]
                vechmod = df.iloc[14,1]
                reg = df.iloc[12,1]
                df.to_csv(claimname+".csv")
outlook = win32com.client.DispatchWithEvents("Outlook.Application", Handler_Class)

#and then an infinit loop that waits from events.
pythoncom.PumpMessages()