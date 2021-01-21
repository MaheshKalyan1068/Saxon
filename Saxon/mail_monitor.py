# -*- coding: utf-8 -*-
"""
Created on Sun Jul 19 01:58:35 2020

@author: Manomay
"""
import html5lib
import win32com.client
import pythoncom
from pywinauto.keyboard import send_keys
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
import urllib
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from python_imagesearch.imagesearch import imagesearch
import pywinauto


class Handler_Class(object):
    def OnNewMailEx(self, receivedItemsIDs):
        # RecrivedItemIDs is a collection of mail IDs separated by a ",".
        # You know, sometimes more than 1 mail is received at the same moment.
        try:
            for ID in receivedItemsIDs.split(","):
                mail = outlook.Session.GetItemFromID(ID)
                subject = mail.Subject
                if subject == "Claims - Claims Form (FNOL)":
                    #cnxn = pyodbc.connect('DRIVER={SQL Server};server=183.82.0.186;PORT=1433;database=Saxon_Demo;uid=MANOMAY1;pwd=manomay1')
                    cnxn = pyodbc.connect('DRIVER={SQL Server};server=192.168.2.150;database=Saxon_Demo;uid=MANOMAY1;pwd=Manomay1')
                    #print("yes")
                    
                    msg = mail.Body
                    #print(msg)
                    msg1 = str(msg)
                    #print(type(msg1))
                    #msg_1=msg1.replace("\n","")
                    msg_1= "".join([s for s in msg1.strip().splitlines(True) if s.strip()])
                    kt = msg_1.replace(r"\n"," ")
                    print(kt)
                    kt1 =''.join(kt.splitlines())
                    print(kt1 ,"KTKTKTKTKTKTKTKTKTK")
                    #print (" ".join(line.strip() for line in msg_1)  )
                    #print(type(msg_1))
                    #print(type(msg_1,"typeof msg_1"))
                    #print(msg_1,"?????")
                    #print(type(msg_1),"?????????????????????????????")
                    #print()
                    #msign = re.search("View Signature <"+"(.+?)"+"dl=0>",kt1)
                    #print(msign,"MSIGNMSDIGN")
                    #signreg= msign.group(1)
                    #print(str(kt),"--------------------------------------------------------")
                    kt2 = kt1.replace("    "," ")
                    print(kt2)
                    #print(signreg)
                    receviedtime = mail.ReceivedTime
                    #print(receviedtime)
                    temp = str(receviedtime).rstrip("+00:00").strip()
                    #print(temp)
                    y= datetime.datetime.strptime(temp,"%Y-%m-%d %H:%M:%S")
                    print(y)
                    p =y.strftime('%Y-%m-%d %H:%M:%S')
                    #print(y)
                    #print(p)
                    msghtml = mail.HTMLBody
                    #print(type(msghtml))
                    #print(msghtml)
                    #msgbody = mail.Body
                    #print(msgbody)
                    #df18 = pd.read_html(mail.HTMLBody)
                    #print(df18,"kokoko")
                    dfs = pd.read_html(msghtml,flavor='html5lib')
                    #print(type(dfs))
                    #print(dfs,"dfssss.csv")
                    df0=dfs[0]
                    df0.to_csv("df000.csv")
                    #print("&"*50)
                    #print(df0.iloc[5,1])
                    #print("%"*50)
                    drivers = df0.iloc[2,1]
                    df1=dfs[2]
                    df1.to_csv("df111.csv")
                    df = df1.replace(np.nan, '', regex=True)
                    #print(df)
                    df.to_csv("tollgate.csv")
                    date_time = df.iloc[17,1]
                    #datest = date_time.split()
                    #datef = dateutil.parser.parse(date_time)
                    t = dateparser.parse(date_time)
                    #oldformat = datef
                    #datetimeobject = datetime.datetime.strptime(str(oldformat),'%Y-%m-%d  %H:%M:%S')
                    newformat = t.strftime('%d/%m/%Y')
                    #print (newformat)
                    newtime = t.strftime('%I:%M %p')
                    #print(type(newtime))
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
                    print(newformat,"/n",newtime,"/n",ad,"/n",street,"/n",zipco,"/n",reg,"/n",vechmaker)
                    sleep(1.5)
                    dateint1 = datetime.datetime.now()
                    dateint = dateint1.strftime("%d/%m/%Y %H:%M:%S")
                    sql = "INSERT INTO Saxon_Claims(Formstack_Link, Reported_By,PolicyNumber, Bot_Execution_Status, Bot_Start_Time, Bot_End_Time,ClaimNumber, Error_Details, Actions) VALUES (?,?,?,?,?,?,?,?,?)"
                    val = (p,claimname,policy_number,"InProgress",dateint1," "," "," "," ")
                    cursor = cnxn.cursor()
                    cursor.execute(sql, val)
                    cnxn.commit()
                    chrome = webdriver.Chrome(executable_path='C:\Program Files\chromedriver.exe')
                    #chrome = webdriver.Chrome()
                    
                    chrome.get("https://qasims.saxon.ky/Sims/(S(0vciitns3ckae0h4egjnq2ik))/Login.aspx")
                    chrome.maximize_window()
                    
                    # wait for element to appear, then hover it
                    #wait = WebDriverWait(chrome, 10)
                    uname = WebDriverWait(chrome, 10).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#Login1_UserName"))))
                    uname.send_keys("Priya")
                    pwd = chrome.find_element_by_css_selector("#Login1_Password")
                    pwd.send_keys("Saxon2018!")
                    chrome.find_element_by_css_selector("#Login1_Button1").click()
                    sleep(2.5)
                    New_claim = WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl00_btnNewClaim"))))
                    New_claim.click()
                    
                    sleep(4)
                    chrome.switch_to.frame("GridEditor")
                    
                    sleep(3)
                    
                    chrome.find_element_by_id("ctl03_ctl00_wizNewClaim_chkCoverageVerification").click()
                    #check1 = WebDriverWait(chrome, ).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl03_ctl00_wizNewClaim_chkCoverageVerification"))))
                    #check1.click()
                    
                    loss_date=chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_rdpIncidentDate_dateInput")
                    loss_date.send_keys(newformat)
                    
                    Policy_Picker = chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_sppPolicy_lnkPolicyPicker").click()
                    chrome.switch_to.frame("rwPolicyPicker")
                    sleep(3)
                    
                    inputElement = chrome.find_element_by_name("tbPolicyNumber")
                    inputElement.click()
                    inputElement.send_keys(policy_number)
                    #Polinpt = WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.NAME,("#tbPolicyNumber"))))
                    #Polinpt.click()
                    #Polinpt.send_keys("P05062018-KY-051268")
                    
                    chrome.find_element_by_css_selector("#btnOK").click()
                    sleep(3.5)
                    print("vali1")
                    btnselect = chrome.find_elements_by_css_selector("#rgPolicy_ctl00_ctl04_btnSelect")
                    print(chrome.window_handles,"handlessssss")
                    if not btnselect:
                        print("hi")
                        sleep(1.5)
                        chrome.switch_to.default_content()
                        sleep(0.5)
                        chrome.switch_to.frame("GridEditor")
                        sleep(1)
                        chrome.find_element_by_css_selector("#RadWindowWrapper_ctl03_ctl00_rwPolicyPicker > div.rwTitleBar > div > ul > li:nth-child(5) > span").click()
                        chrome.switch_to.default_content()
                        sleep(1)
                        chrome.find_element_by_css_selector("#RadWindowWrapper_ctl00_GridEditor > div.rwTitleBar > div > ul > li:nth-child(6) > span").click()
                        sleep(5)
                        chrome.execute_script("scroll(250, 0)")
                        sleep(1.5)
                        pos = imagesearch("logout.JPG")
                        if pos[0] != -1:
                            pass
                        else:
                            print("page not loaded")
                        sleep(1)
                        pywinauto.mouse.click(button='left', coords=pos)
                        sleep(0.5)
                        chrome.find_element_by_class_name("rwOkBtn").click()
                        sleep(2)
                        chrome.close()
                    btnselect = WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#rgPolicy_ctl00_ctl04_btnSelect"))))
                    btnselect.click()
                    #chrome.find_element_by_css_selector("#rgPolicy_ctl00_ctl04_btnSelect").click()
                    chrome.switch_to.default_content()
                    #chrome.switch_to.frame("contents")
                    chrome.switch_to.frame("GridEditor")
                    sleep(2.5)
                    #btnselect = WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl03_ctl00_wizNewClaim_chkUseInsured"))))
                    
                    chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_chkUseInsured").click()
                    #check2 = WebDriverWait(chrome, 10).until(ec.element_to_be_clickable((By.CSS_SELECTOR,("#ctl03_ctl00_wizNewClaim_chkUseInsured"))))##ctl03_ctl00_wizNewClaim_chkUseInsured
                    #check2.click()
                    
                    step2nxt = WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl03_ctl00_wizNewClaim_StartNavigationTemplateContainerID_btnNext"))))
                    step2nxt.click()
                    #chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_StartNavigationTemplateContainerID_btnNext").click()
                    sleep(2)
                    #emailid
                    emafrma = df.iloc[6,1]
                    if emafrma != "NaN":
                        emailidma = WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl03_ctl00_wizNewClaim_txtClaimantEmail"))))
                        emailidma.send_keys(emafrma)
                        sleep(0.7)
                    #phne =df.iloc[5,1]
                    """if phne != "NaN": 
                        t = chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_txtClaimantPhone")
                        sleep(0.7)
                        t.clear()
                        sleep(0.7)
                        t.click()
                        sleep(0.7)
                        t.send_keys(phne)"""
                    step2button = WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl03_ctl00_wizNewClaim_StepNavigationTemplateContainerID_Button1"))))
                    step2button.click()
                    #chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_StepNavigationTemplateContainerID_Button1").click()
                    sleep(4)
                    okbutton = WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CLASS_NAME,("rwOkBtn"))))
                    okbutton.click()
                    #chrome.find_element_by_class_name("rwOkBtn").click()
                    
                    time_picker = chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_txtInjuryTime_dateInput")
                    time_picker.send_keys(newtime)
                    secondary_loss = chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_rdpIncidentDateSecondary_dateInput")
                    secondary_loss.send_keys(newformat)
                    
                    loss_location = chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_addrLossLocationAddress_lnkAddress").click()
                    ad1 = chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_addrLossLocationAddress_txtAddress1")
                    ad1.send_keys(street)
                    sleep(0.5)
                    text = chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_addrLossLocationAddress_cboZipCitySearch_Input")
                    text.send_keys(inti)
                    sleep(5)
                    btnselect = WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl03_ctl00_wizNewClaim_addrLossLocationAddress_cboZipCitySearch_DropDown > div > ul > li.rcbHovered"))))
                    
                    
                    textlist = chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_addrLossLocationAddress_cboZipCitySearch_DropDown > div")
                    li = textlist.parent.find_elements_by_tag_name('li')
                    text_str = textlist.text
                    textsplit = text_str.split("\n")
                    clikindex = textsplit.index(zipco)
                    li[clikindex].click()
                    chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_addrLossLocationAddress_lnkAddress").click()
                    
                    examin = chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_cboExaminer_Input")
                    examin.send_keys("priya")
                    #claimant = chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_txtInjuryDescription")
                    #claimant.send_keys("Vechile Damaged")
                    desofacc = df.iloc[19,1]
                    if desofacc != "NaN":
                        claimloss = chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_txtClaimLossDesc")
                        claimloss.send_keys(desofacc)
                    step3_next = chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_StepNavigationTemplateContainerID_Button1").click()
                    sleep(6)
                    btnselect = WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl03_ctl00_wizNewClaim_cboInsuredVehicle_Input"))))
                    vno = chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_cboInsuredVehicle_Input")
                    
                    op = chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_cboInsuredVehicle_Arrow").click()
                    sleep(0.5)
                    li = vno.parent.find_elements_by_tag_name('li')
                    li[1].click()
                    sleep(2)
                    #v1 = chrome.find_element_by_css_selector("ctl03_ctl00_wizNewClaim_cboInsuredVehicle_Input").click()
                    #sleep(2)
                    intvecdes = df.iloc[21,1]
                    
                    
                    sleep(5.5)
                    if intvecdes != "NaN":
                        intvechicledes = chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_txtInsuredVehicleDesc")
                        intvechicledes.send_keys(intvecdes)
                    #namdd1 = chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_cboInsuredDriver_Arrow")
                    ##ctl03_ctl00_wizNewClaim_cboInsuredDriver_DropDown > div > ul > li.rcbHovered
                    if df.iloc[0,1] == "Yes":
                        print("KALACHASHMA")
                        WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl03_ctl00_wizNewClaim_cboInsuredDriver_Arrow")))).click()
                        sleep(1)
                        lst = WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl03_ctl00_wizNewClaim_cboInsuredDriver_DropDown"))))
                        ni = lst.parent.find_elements_by_tag_name('li')
                        ni[1].click()
                    else:
                        WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl03_ctl00_wizNewClaim_cboInsuredDriver_Arrow")))).click()
                        sleep(1)
                        lst = WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl03_ctl00_wizNewClaim_cboInsuredDriver_DropDown"))))
                        ni = lst.parent.find_elements_by_tag_name('li')
                        ni[0].click()
                        driname = df.iloc[8,1]
                        lstdriname = driname.split(" ")
                        fn=chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_txtInsuredDriverFirstName")
                        fn.send_keys(lstdriname[0])
                        ln=chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_txtInsuredDriverLastName")
                        ln.send_keys(lstdriname[1])
                        
                    sleep(3)
                    #if df.iloc[] != 
                    typevech = chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_cboInsuredVehicleStyle_Input").get_property("value")
                    vech = WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl03_ctl00_wizNewClaim_cboClaimantVehicleType_Input"))))
                    vech.send_keys(typevech)
                    sleep(0.25)
                    plate =WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl03_ctl00_wizNewClaim_txtClaimantVehicleLicensePlate"))))
                    plate.send_keys(reg)
                    sleep(0.25)
                    yearint = chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_txtInsuredVehicleYear").get_property("value")
                    year = WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl03_ctl00_wizNewClaim_txtClaimantVehicleYear"))))
                    year.send_keys(yearint)
                    sleep(0.25)
                    islandtxt = chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_cboInsuredVehicleState_Input").get_property("value")
                    island = WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl03_ctl00_wizNewClaim_cboClaimantVehicleState_Input"))))
                    island.send_keys(islandtxt)
                    sleep(0.25)
                    make = WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl03_ctl00_wizNewClaim_cboClaimantVehicleMake_Input"))))
                    make.send_keys(vechmaker)
                    sleep(0.25)
                    model = WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl03_ctl00_wizNewClaim_cboClaimantVehicleModel_Input"))))
                    model.send_keys(vechmod)
                    sleep(0.25)
                    colourtxt = chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_cboInsuredVehicleColor_Input").get_property("value")
                    print(colourtxt)
                    colour = WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl03_ctl00_wizNewClaim_cboClaimantVehicleColor_Input"))))
                    colour.send_keys(colourtxt)
                    sleep(0.25)
                    othvecdesma = df.iloc[23,1]
                    if othvecdesma!="NaN":
                        othvecdes = WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl03_ctl00_wizNewClaim_txtClaimantVehicleDesc"))))
                        othvecdes.send_keys(othvecdesma)
                    licentext = chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_txtInsuredDriverLicenseNumber").get_property("value")
                    licen = WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl03_ctl00_wizNewClaim_txtDriverLicenseNumber"))))
                    licen.send_keys(licentext)
                    sleep(1)
                    coverage = chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_cboCoverage_Input").click()
                    sleep(1.5)
                    lbox = chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_cboCoverage_DropDown > div.rcbScroll.rcbWidth")
                    li = lbox.parent.find_elements_by_tag_name('li')
                    c1 = chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_cboCoverage_i0_CheckBox").click()
                    sleep(0.4)
                    #c2 = chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_cboCoverage_i4_CheckBox").click()
                    #sleep(0.4)
                    #c3= chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_cboCoverage_i1_CheckBox").click()
                    
                    #sleep(0.4)
                    losstype = chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_cboLossType_Input")
                    losstype.send_keys("Auto")
                    sleep(3)
                    chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_cboInsuredVehicle_Input").click()
                    sleep(2)
                    chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_FinishNavigationTemplateContainerID_btnCreate").click()
                    sleep(8)
                    WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CLASS_NAME,("rwDialogMessage"))))
                    t=chrome.find_element_by_class_name("rwDialogMessage")
                    poltext = t.text
                    res = ''.join(filter(lambda i: i.isdigit(), poltext))
                    claimno = "#"+res
                    print(claimno)
                    sleep(2)
                    chrome.find_element_by_class_name("rwOkBtn").click()
                    sleep(6)
                    inbox = win32com.client.gencache.EnsureDispatch("Outlook.Application").GetNamespace("MAPI")
                    #print(dir(inbox))
                    inbox = win32com.client.Dispatch("Outlook.Application")
                    #print(dir(inbox))
                    mail = inbox.CreateItem(0x0)
                    #mail.To = "mahesh.kalyankar@manomay.biz"
                    mail.To = "deepika.bharatula@manomay.biz"
                    m = mail.To
                    #mail.Attachments.Add(r"C:\Users\Manomay\\"+policy_number+'.png')
                    #mail.CC = "testcc@test.com"
                    mail.Subject = "Claim Logged mail"
                    mail.Body = "Hi " +claimname +",\n"+"claim as logged as per your request"+"\n"+"Please find the claim numnver: "+claimno  +"\n"+"\n" +"Thank you."+"\n"+"Saxon Claims Bot."
                                    
                    mail.Send()
                    sleep(1)
                    #chrome.switch_to.window("GridEditor")
                    #sleep(3)
                    #chrome.switch_to.window(chrome.window_handles[1])
                    chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_StartNavigationTemplateContainerID_btnCancel").click()
                    sleep(10)
                    OPN = df.iloc[31,1]
                    OPPN = df.iloc[32,1]
                    OPEM = df.iloc[33,1]
                    OPIC = df.iloc[34,1]
                    COPPRE = df.iloc[26,1]
                    TOWED = df.iloc[25,1]
                    #OPPN = df.iloc[36,1]
                    ##Other claimet Name
                    print(chrome.window_handles,"handlessssss")
                    chrome.switch_to.window(chrome.window_handles[0])
                    sleep(3)
                    print(OPN,"OPN")
                    print(OPPN,"OPPN")
                    print(OPEM,"OPEM")
                    print(OPIC,"OPIC")
                    print(COPPRE,"COPPRE")
                    print(TOWED,"TOWED")
                    print(OPN,OPPN,OPEM,OPIC,COPPRE,TOWED)
                    if COPPRE != "" or TOWED!= "":
                        chrome.find_element_by_css_selector("#ctl00_RadMenu1 > ul > li.rmItem.rmLast.recent-claims > span").click()
                        sleep(1.5)
                        WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl00_RadMenu1 > ul > li.rmItem.rmLast.rmSelected.recent-claims > div > ul > li.rmItem.rmFirst")))).click()
                        sleep(6)
                        WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl00_lblCurrentClaimModule")))).click()
                        sleep(2)
                        print("claimmanitance")
                        WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl00_rptClaimModules_ctl01_divModuleName")))).click()
                        sleep(5)
                        print("selecting")
                        ADdate =WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl00_MainContent_rdpCAInformedDate_dateInput"))))
                        ADdate.send_keys("27/12/2020")
                        if COPPRE != "" :
                            ad = chrome.find_element_by_css_selector("#ctl00_MainContent_cfClaim1_cboCustome074c2a2c42d4d5fa7434b921e08158e_Input")
                            sleep(1.5)
                            ad.send_keys(COPPRE)
                            chrome.execute_script("scroll(250, 0)")
                        if TOWED!= "Yes":
                            chrome.execute_script("scroll(250, 0)")
                            sleep(0.5)
                            pos = imagesearch("towed.JPG")
                            if pos[0] != -1:
                                pass
                            else:
                                print("page not loaded")
                            sleep(1)
                            pywinauto.mouse.click(button='left', coords=pos)
                            sleep(12)
                            chrome.find_element_by_css_selector("#ctl00_MainContent_chkTowed").click()
                        sleep(9)
                        WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl00_save")))).click()
                        sleep(15)
                        print("saving towed")
                        chrome.find_element_by_class_name("rwOkBtn").click()
                        sleep(2)                    ##Other claimet Phine
                    
                    print("docu intial")
                    Pleasesign = df.iloc[36,1]
                    Driverlicense = df.iloc[39,1]
                    #print(msg_1)
                    #print(Pleasesign,Driverlicense)
                    if df.iloc[36,1] != "":
                        dateofsig = df.iloc[37,1]
                    if Pleasesign != "":
                        msign = re.search("View Signature <"+"(.+?)"+"dl=0>",kt)
                        #print(msign)
                        msignval = msign.group(1)
                        #print(msignval)
                        url_sign = str(msignval)+"dl=1"
                        urllib.request.urlretrieve(url_sign, r'C:\Users\MANOMAY\Desktop\Saxon\Signature.jpg')
                        sleep(2)
                        #chrome.find_element_by_css_selector("#ctl00_RadMenu1 > ul > li.rmItem.rmLast.recent-claims > span").click()
                        #chrome.find_element_by_css_selector("#ctl00_RadMenu1 > ul > li.rmItem.rmLast.rmSelected.recent-claims > div > ul > li.rmItem.rmFirst").click()
                        print("sign1")
                        WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl00_lblCurrentClaimModule")))).click()
                        sleep(3)
                        print("sign2")
                        WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl00_rptClaimModules_ctl12_divModuleName")))).click()
                        print("sign3")
                        sleep(10)
                        WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl00_add")))).click()
                        print("sign4")
                        sleep(3)
                        chrome.switch_to.frame("GridEditor")
                        sleep(2)
                        tit =chrome.find_element_by_css_selector("#ctl03_txtTitle")
                        tit.send_keys("Signature")
                        typeattch =chrome.find_element_by_css_selector("#ctl03_cboAttachmentType_Input")
                        typeattch.send_keys("Image")
                        DLdes = chrome.find_element_by_css_selector("#ctl03_txtDescription")
                        DLdes.send_keys("Signature Uploaded")                                            
                        if df.iloc[37,1] != "":
                            dos = chrome.find_element_by_css_selector("#ctl03_rdpDocumentDate_dateInput")
                            dos.send_keys(dateofsig)
                        print("Sign5")
                        chrome.find_element_by_css_selector("#ctl03_radUploadListContainer").click()
                        print("sign6")
                        send_keys(r'C:\Users\MANOMAY\Desktop\Saxon\Signature.jpg')
                        sleep(2.5)
                        send_keys('{TAB}')
                        sleep(0.5)
                        send_keys('{TAB}')
                        sleep(0.5)
                        send_keys('{ENTER}')
                        sleep(1.25)
                        WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl03_btnSubmit")))).click()
                        sleep(3)
                        chrome.find_element_by_class_name("rwOkBtn").click()
                        sleep(4)
                        #chrome.find_element_by_css_selector("#alert1609398767368_content > div > div.rwDialogButtons > input").click()
                        #sleep(4)
                        
                    
                    if Driverlicense != "":
                        mdrivlic = re.search("Driver's License: View File <"+"(.+?)"+"dl=0>",kt2)
                        mdrivlicval = mdrivlic.group(1)
                        print(mdrivlicval)
                        url_drivlic = str(mdrivlicval)+"dl=1"
                        print("Connecting...")
                        urllib.request.urlretrieve(url_drivlic, r'C:\Users\MANOMAY\Desktop\Saxon\Driverlicense.jpg')
                        print("Saving...")
                        #chrome.find_element_by_css_selector("#ctl00_RadMenu1 > ul > li.rmItem.rmLast.recent-claims > span").click()
                        #chrome.find_element_by_css_selector("#ctl00_RadMenu1 > ul > li.rmItem.rmLast.rmSelected.recent-claims > div > ul > li.rmItem.rmFirst").click()
                        #chrome.find_element_by_css_selector("#ctl00_lblCurrentClaimModule").click()
                        print("moduledropbox")
                        sleep(4)
                        elem = chrome.find_element_by_css_selector("#ctl00_add")
                        if elem.is_displayed():
                            elem.click()
                            #WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl00_rptClaimModules_ctl12_divModuleName")))).click()
                            #WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl00_add")))).click()
                            sleep(3)
                            print("lic1")
                            chrome.switch_to.frame("GridEditor")
                            sleep(3)
                            tit =chrome.find_element_by_css_selector("#ctl03_txtTitle")
                            tit.send_keys("Driver's License")
                            typeattch =chrome.find_element_by_css_selector("#ctl03_cboAttachmentType_Input")
                            typeattch.send_keys("Driver's License")
                            DLdes = chrome.find_element_by_css_selector("#ctl03_txtDescription")
                            DLdes.send_keys("DRver's License Uploaded")
                            print("lic2")
                            #WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl03_radUploadrow2")))).click()
                            chrome.find_element_by_css_selector("#ctl03_radUploadListContainer").click()
                            print("lic3")
                            sleep(0.7)
                            send_keys(r'C:\Users\MANOMAY\Desktop\Saxon\Driverlicense.jpg')
                            sleep(2.5)
                            send_keys('{TAB}')
                            sleep(0.5)
                            send_keys('{TAB}')
                            sleep(0.5)
                            send_keys('{ENTER}')
                            sleep(1.5)
                            WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl03_btnSubmit")))).click()
                            sleep(3)
                            chrome.find_element_by_class_name("rwOkBtn").click()
                            sleep(4)
                            #sleep(3)
                        else:
                            WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl00_lblCurrentClaimModule")))).click()
                            sleep(3)
                            print("sign2")
                            WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl00_rptClaimModules_ctl12_divModuleName")))).click()
                            print("sign3")
                            WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl00_add")))).click()
                            print("est1")
                            sleep(4)
                            chrome.switch_to.frame("GridEditor")
                            sleep(3)
                            tit =chrome.find_element_by_css_selector("#ctl03_txtTitle")
                            tit.send_keys("Driver's License")
                            sleep(0.7)
                            typeattch =chrome.find_element_by_css_selector("#ctl03_cboAttachmentType_Input")
                            typeattch.send_keys("Driver's License")
                            sleep(0.7)
                            DLdes = chrome.find_element_by_css_selector("#ctl03_txtDescription")
                            DLdes.send_keys("DRver's License Uploaded")
                            print("lic2")
                            #WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl03_radUploadrow2")))).click()
                            chrome.find_element_by_css_selector("#ctl03_radUploadListContainer").click()
                            print("lic3")
                            sleep(0.7)
                            send_keys(r'C:\Users\MANOMAY\Desktop\Saxon\Driverlicense.jpg')
                            sleep(2.5)
                            send_keys('{TAB}')
                            sleep(0.5)
                            send_keys('{TAB}')
                            sleep(0.5)
                            send_keys('{ENTER}')
                            sleep(1.5)
                            WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl03_btnSubmit")))).click()
                            sleep(3)
                            chrome.find_element_by_class_name("rwOkBtn").click()
                            sleep(4)
                        
                    Repairest = df.iloc[41,1]
                    if Repairest != "":
                        mest = re.search("Repair Estimate:  View File <"+"(.+?)"+"dl=0>",kt2)
                        mestval = mest.group(1)
                        print(mestval)
                        url_repair = mestval+"dl=1"
                        urllib.request.urlretrieve(url_repair, r'C:\Users\MANOMAY\Desktop\Saxon\Repair_Estimate.jpg')
                        #chrome.find_element_by_css_selector("#ctl00_RadMenu1 > ul > li.rmItem.rmLast.recent-claims > span").click()
                        #chrome.find_element_by_css_selector("#ctl00_RadMenu1 > ul > li.rmItem.rmLast.rmSelected.recent-claims > div > ul > li.rmItem.rmFirst").click()
                        #WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl00_lblCurrentClaimModule")))).click()
                        #WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl00_rptClaimModules_ctl12_divModuleName")))).click()
                        elem = WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl00_add"))))
                        if elem.is_displayed():
                            elem.click() # this will click the element if it is there
                            print("FOUND THE LINK CREATE ACTIVITY! and Clicked it!")
                            #WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl00_add")))).click()
                            print("est1")
                            sleep(4)
                            chrome.switch_to.frame("GridEditor")
                            sleep(3)
                            tit =chrome.find_element_by_css_selector("#ctl03_txtTitle")
                            tit.send_keys("Repair Estimates")
                            typeattch =chrome.find_element_by_css_selector("#ctl03_cboAttachmentType_Input")
                            typeattch.send_keys("Estimate")
                            DLdes = chrome.find_element_by_css_selector("#ctl03_txtDescription")
                            DLdes.send_keys("Repair Estimates Uploaded")
                            #WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl03_radUploadrow2")))).click()
                            chrome.find_element_by_css_selector("#ctl03_radUploadListContainer").click()
                            sleep(0.7)
                            send_keys(r"C:\Users\MANOMAY\Desktop\Saxon\Repair_Estimate.jpg")
                            sleep(2.5)
                            send_keys('{TAB}')
                            sleep(0.5)
                            send_keys('{TAB}')
                            sleep(0.5)
                            send_keys('{ENTER}')
                            sleep(1.5)
                            chrome.find_element_by_css_selector("#ctl03_btnSubmit").click()
                            sleep(3)
                            chrome.find_element_by_class_name("rwOkBtn").click()
                            #WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#alert1608665700390_content > div > div.rwDialogButtons > input")))).click()
                        else:
                            WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl00_lblCurrentClaimModule")))).click()
                            sleep(3)
                            print("sign2")
                            WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl00_rptClaimModules_ctl12_divModuleName")))).click()
                            print("sign3")
                            WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl00_add")))).click()
                            print("est1")
                            sleep(4)
                            chrome.switch_to.frame("GridEditor")
                            sleep(2)
                            tit =chrome.find_element_by_css_selector("#ctl03_txtTitle")
                            tit.send_keys("Repair Estimates")
                            typeattch =chrome.find_element_by_css_selector("#ctl03_cboAttachmentType_Input")
                            typeattch.send_keys("Estimate")
                            DLdes = chrome.find_element_by_css_selector("#ctl03_txtDescription")
                            DLdes.send_keys("Repair Estimates Uploaded")
                            #WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl03_radUploadrow2")))).click()
                            chrome.find_element_by_css_selector("#ctl03_radUploadListContainer").click()
                            send_keys(r"C:\Users\MANOMAY\Desktop\Saxon\Repair_Estimate.jpg")
                            sleep(2.5)
                            send_keys('{TAB}')
                            sleep(0.5)
                            send_keys('{TAB}')
                            sleep(0.5)
                            send_keys('{ENTER}')
                            sleep(1.5)
                            chrome.find_element_by_css_selector("#ctl03_btnSubmit").click()
                            sleep(3)
                            chrome.find_element_by_class_name("rwOkBtn").click()
                            sleep(5)
                    sleep(2)
                    
                    
                    if OPN != "" or OPPN != "" or OPEM != "" or OPIC != "":
                        print("POLO")
                        #sleep(1)
                       #polint = chrome.find_element_by_css_selector("#txtClaimSearch").click()
                       #polint.sendkeys(res)
                        #sleep(0.7)
                        
                       #chrome.find_element_by_css_selector("#toolbar > div > div.get-claim > a").click()
                        print("HI")
                        #chrome.execute_script("scroll(250, 0)")
                        #chrome.switch_to.window(chrome.window_handles[0])
                        chrome.find_element_by_css_selector("#ctl00_RadMenu1 > ul > li.rmItem.rmLast.recent-claims > span").click()
                        sleep(1.5)
                        WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl00_RadMenu1 > ul > li.rmItem.rmLast.rmSelected.recent-claims > div > ul > li.rmItem.rmFirst")))).click()
                        sleep(6)
                        WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl00_claimInfo_divCIClaimant")))).click()
                        sleep(1.5)
                        WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl00_claimInfo_lnkAddClaimant")))).click()
                        sleep(4)
                        WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl00_MainContent_rdbIndividual")))).click()
                        
                        sleep(2)
                        ADdate =chrome.find_element_by_css_selector("#ctl00_MainContent_rdpCAInformedDate_dateInput")
                        #from datetime import timedelta
                        #yesterday = datetime.now() - timedelta(1)
                        #adint =datetime.strftime(yesterday, '%d-%m-%Y')
                        ADdate.send_keys("27/12/2020")
                        sleep(1)
                        WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl00_MainContent_txtFullValue")))).click()
                        print("OPN")
                        sleep(1)
                        #sleep(0.8)
                        if OPN != "":
                            print("OPN1")
                            lstnames= OPN.split(" ")
                            fname = chrome.find_element_by_css_selector("#ctl00_MainContent_txtClaimantFirstName")
                            fname.send_keys(lstnames[0])
                            sleep(1)
                            lname = chrome.find_element_by_css_selector("#ctl00_MainContent_txtClaimantLastName")
                            lname.send_keys(lstnames[1])
                        if OPPN != "":
                            print("OPN2")
                            #chrome.find_element_by_css_selector("#ctl00_RadMenu1 > ul > li.rmItem.rmLast.rmSelected.recent-claims > span").click()
                            Phnenum = chrome.find_element_by_css_selector("#ctl00_MainContent_txtMobilePhone")
                            Phnenum.send_keys(OPPN)
                             
                        if OPEM != "":
                            print("OPN3")
                            #chrome.find_element_by_css_selector("#ctl00_RadMenu1 > ul > li.rmItem.rmLast.rmSelected.recent-claims > span").click()
                            emailadd = chrome.find_element_by_css_selector("#ctl00_MainContent_txtEmail")
                            emailadd.send_keys(OPEM)
                            
                    
                        if OPIC != "":
                            print("OPN4")
                            #chrome.find_element_by_css_selector("#ctl00_RadMenu1 > ul > li.rmItem.rmLast.rmSelected.recent-claims > span").click()
                            Incom = chrome.find_element_by_css_selector("#ctl00_MainContent_txtCompanyName")
                            Incom.send_keys(OPEM)
                        chrome.execute_script("scroll(250, 0)")
                        sleep(3)
                        print("Baga")
                        chrome.find_element_by_css_selector("#ctl00_save").click()
                        print("ARAMBOL")
                        sleep(15)
                        WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CLASS_NAME,("rwOkBtn")))).click()
                        print("Khali")
                        sleep(8)
                    #chrome.find_element_by_css_selector("#system-nav > div.nav-toggle.no-select").click()
                    #sleep(1.5)
                    #chrome.find_element_by_css_selector("#ctl00_RadMenu2 > ul > li.rmItem.rmTemplate.logout-nav > div > i").click()
                    chrome.execute_script("scroll(250, 0)")
                    sleep(0.5)
                    pos = imagesearch("logout.JPG")
                    if pos[0] != -1:
                        pass
                    else:
                        print("page not loaded")
                    sleep(1)
                    pywinauto.mouse.click(button='left', coords=pos)
                    sleep(0.5)
                    chrome.find_element_by_class_name("rwOkBtn").click()
                    sleep(2)
                    chrome.close()
                        #WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#alert1609152655887_content > div > div.rwDialogButtons > input")))).click()
                        #chrome.find_element_by_css_selector("#alert1609152655887_content > div > div.rwDialogButtons > input").click()
                        
                    
                    #
                    #chrome.close()
                    #sleep(5)"""
                    submittime1 = datetime.datetime.now()
                    submittime = submittime1.strftime("%d/%m/%Y %H:%M:%S")
                    #chrome.close()
                    #cursor = cnxn.cursor()"""
                    #sql = "INSERT INTO Saxon_Claims(Formstack_Link, DateTime_Submitted, Name, PolicyNumber, BotExecution_Status, DateTime_of_Input, ClaimNumber, Details, Actions, Remarks, No_Of_Runs) VALUES (?,?,?,?,?,?,?,?,?,?,?)"
                    
                    #sql = "INSERT INTO Saxon_Claims(Formstack_Link, Reported_By,PolicyNumber, Bot_Execution_Status, Bot_Start_Time, Bot_End_Time,ClaimNumber, Error_Details, Actions) VALUES (?,?,?,?,?,?,?,?,?)"
                    #val = (p,claimname,policy_number,"Success",dateint,submittime1,claimno," "," ")
                    #print(val)
                    #cursor.execute(sql, val)
                    #cnxn.commit()
                    #cursor.close()
                    #cnxn.close()
                    sql = "UPDATE Saxon_Claims SET Bot_Execution_Status = ?, Bot_End_Time = ?,ClaimNumber =? WHERE ID = (SELECT max(ID) FROM Saxon_Claims)"
                    val = ("Success",submittime1,claimno)
                    cursor.execute(sql, val)
                    cursor.commit()
                    cursor.close()
                    cnxn.close()
        except Exception as e:
            print(e)
            t=chrome.get_log('browser')
            #print(t)
            #print(e)
            chrome.save_screenshot(r"C:\Users\Manomay\\"+policy_number+'.png')
            #chrome.close()
            inbox = win32com.client.gencache.EnsureDispatch("Outlook.Application").GetNamespace("MAPI")
            #print(dir(inbox))
            inbox = win32com.client.Dispatch("Outlook.Application")
            #print(dir(inbox))
            mail = inbox.CreateItem(0x0)
            #mail.To = "mahesh.kalyankar@manomay.biz"
            mail.To = "deepika.bharatula@manomay.biz"
            m = mail.To
            mail.Attachments.Add(r"C:\Users\Manomay\\"+policy_number+'.png')
            #mail.CC = "testcc@test.com"
            mail.Subject = "Saxon Bot Failure Report"
            mail.Body = "Hi " + m.split(".")[0].capitalize() +"\n" +"Claim Submission for "+policy_number+" reported by "+claimname+" has failed."+"\n"+"Please find attached the error message/screenshot"+"\n" +"Errored occured due to unable to map datat to textbox"+"\n"+str(e)+"\n" +"Thank you."+"\n"+"Claims Bot."
                            
            mail.Send()
            #cnxn = pyodbc.connect('DRIVER={SQL Server};server=183.82.0.186;PORT=1433;database=Saxon_Demo;uid=MANOMAY1;pwd=manomay1')
            cnxn = pyodbc.connect('DRIVER={SQL Server};server=192.168.2.150;database=Saxon_Demo;uid=MANOMAY1;pwd=Manomay1')
            cursor = cnxn.cursor()
            submittime1 = datetime.datetime.now()
            submittime = submittime1.strftime("%d/%m/%Y %H:%M:%S")
            #sql = "INSERT INTO Saxon_Claims(Formstack_Link, DateTime_Submitted, Name, PolicyNumber, BotExecution_Status, DateTime_of_Input, ClaimNumber, Details, Actions, Remarks, No_Of_Runs) VALUES (?,?,?,?,?,?,?,?,?,?,?)"
            """sql = "INSERT INTO Saxon_Claims(Formstack_Link, Reported_By,PolicyNumber, Bot_Execution_Status, Bot_Start_Time, Bot_End_Time,ClaimNumber, Error_Details, Actions) VALUES (?,?,?,?,?,?,?,?,?)"
            val = (p,claimname,policy_number,"Failure",dateint,submittime," ","Error Occured","RERUN")
            print(val)
            cursor.execute(sql, val)
            cnxn.commit()
            cursor.close()
            cnxn.close()"""
            sql = "UPDATE Saxon_Claims SET Bot_Execution_Status = ?, Bot_End_Time = ?,ClaimNumber =? ,Error_Details = ?,Actions = ? WHERE ID = (SELECT max(ID) FROM Saxon_Claims)"
            val = ("Failure",submittime1," ","Error Occured While execution","RERUN")
            cursor.execute(sql, val)
            cursor.commit()
            cursor.close()
            cnxn.close()

outlook = win32com.client.DispatchWithEvents("Outlook.Application", Handler_Class)

#and then an infinit loop that waits from events.
pythoncom.PumpMessages()