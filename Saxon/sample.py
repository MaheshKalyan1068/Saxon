# -*- coding: utf-8 -*-
"""
Created on Fri Jul 17 23:29:24 2020

@author: Manomay
"""


from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from time import sleep


chrome = webdriver.Chrome()

chrome.get("https://qasims.saxon.ky/Sims/(S(0vciitns3ckae0h4egjnq2ik))/Login.aspx")
#chrome.maximize_window()

# wait for element to appear, then hover it
wait = WebDriverWait(chrome, 10)
uname = WebDriverWait(chrome, 10).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#Login1_UserName"))))
uname.send_keys("Priya")
pwd = chrome.find_element_by_css_selector("#Login1_Password")
pwd.send_keys("Saxon2018!")
chrome.find_element_by_css_selector("#Login1_Button1").click()

New_claim = WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl00_btnNewClaim"))))
New_claim.click()

sleep(1)
chrome.switch_to.frame("GridEditor")

sleep(3)

chrome.find_element_by_id("ctl03_ctl00_wizNewClaim_chkCoverageVerification").click()
#check1 = WebDriverWait(chrome, ).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl03_ctl00_wizNewClaim_chkCoverageVerification"))))
#check1.click()

loss_date=chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_rdpIncidentDate_dateInput")
loss_date.send_keys("Jun 14, 2020")

Policy_Picker = chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_sppPolicy_lnkPolicyPicker").click()
chrome.switch_to.frame("rwPolicyPicker")
sleep(2)

inputElement = chrome.find_element_by_name("tbPolicyNumber")
inputElement.click()
inputElement.send_keys("P05062018-KY-051268")
#Polinpt = WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.NAME,("#tbPolicyNumber"))))
#Polinpt.click()
#Polinpt.send_keys("P05062018-KY-051268")

chrome.find_element_by_css_selector("#btnOK").click()
sleep(1)

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
step2button = WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl03_ctl00_wizNewClaim_StepNavigationTemplateContainerID_Button1"))))
step2button.click()
#chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_StepNavigationTemplateContainerID_Button1").click()
sleep(4)
okbutton = WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CLASS_NAME,("rwOkBtn"))))
okbutton.click()
#chrome.find_element_by_class_name("rwOkBtn").click()

time_picker = chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_txtInjuryTime_dateInput")
time_picker.send_keys("1:00 PM")
secondary_loss = chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_rdpIncidentDateSecondary_dateInput")
secondary_loss.send_keys("Jun 14, 2020")

loss_location = chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_addrLossLocationAddress_lnkAddress").click()
ad1 = chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_addrLossLocationAddress_txtAddress1")
ad1.send_keys("Street 305")
sleep(0.5)
text = chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_addrLossLocationAddress_cboZipCitySearch_Input")
text.send_keys("General")
sleep(5)
btnselect = WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl03_ctl00_wizNewClaim_addrLossLocationAddress_cboZipCitySearch_DropDown > div > ul > li.rcbHovered"))))


textlist = chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_addrLossLocationAddress_cboZipCitySearch_DropDown > div")
li = textlist.parent.find_elements_by_tag_name('li')
text_str = textlist.text
textsplit = text_str.split("\n")
clikindex = textsplit.index("General, KY 01100")
li[clikindex].click()
chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_addrLossLocationAddress_lnkAddress").click()

examin = chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_cboExaminer_Input")
examin.send_keys("priya")
claimant = chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_txtInjuryDescription")
claimant.send_keys("Vechile Damaged")
claimloss = chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_txtClaimLossDesc")
claimloss.send_keys("Accident")
step3_next = chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_StepNavigationTemplateContainerID_Button1").click()
sleep(6)
btnselect = WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl03_ctl00_wizNewClaim_cboInsuredVehicle_Input"))))
vno = chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_cboInsuredVehicle_Input")

op = chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_cboInsuredVehicle_Arrow").click()
sleep(0.5)
li = vno.parent.find_elements_by_tag_name('li')
li[1].click()
sleep(1)
dri =WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl03_ctl00_wizNewClaim_cboInsuredDriver_Input"))))
sleep(1.5)
#namdd1 = chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_cboInsuredDriver_Arrow")
namddlst = WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl03_ctl00_wizNewClaim_cboInsuredDriver > span > button")))).click()
sleep(1.5)
lst = chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_cboInsuredDriver_DropDown")
ni = lst.parent.find_elements_by_tag_name('li')
ni[1].click()
    
sleep(3)
vech = WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl03_ctl00_wizNewClaim_cboClaimantVehicleType_Input"))))
vech.send_keys("car")
sleep(0.25)
plate =WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl03_ctl00_wizNewClaim_txtClaimantVehicleLicensePlate"))))
plate.send_keys("HGF54")
sleep(0.25)
year = WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl03_ctl00_wizNewClaim_txtClaimantVehicleYear"))))
year.send_keys("2015")
sleep(0.25)
island = WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl03_ctl00_wizNewClaim_cboClaimantVehicleState_Input"))))
island.send_keys("CAYMAN ISLANDS")
sleep(0.25)
make = WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl03_ctl00_wizNewClaim_cboClaimantVehicleMake_Input"))))
make.send_keys("Toyota")
sleep(0.25)
model = WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl03_ctl00_wizNewClaim_cboClaimantVehicleModel_Input"))))
model.send_keys("AVALON")
sleep(0.25)
colour = WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl03_ctl00_wizNewClaim_cboClaimantVehicleColor_Input"))))
colour.send_keys("Grey")
sleep(0.25)
licen = WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl03_ctl00_wizNewClaim_txtDriverLicenseNumber"))))
licen.send_keys("HGJH56465")
sleep(1)
coverage = chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_cboCoverage_Input").click()
sleep(1.5)
lbox = chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_cboCoverage_DropDown > div.rcbScroll.rcbWidth")
li = lbox.parent.find_elements_by_tag_name('li')
c1 = chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_cboCoverage_i0_CheckBox").click()
sleep(0.2)
c2 = chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_cboCoverage_i4_CheckBox").click()
sleep(0.2)
c3= chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_cboCoverage_i7_CheckBox").click()

sleep(0.2)
losstype = chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_cboLossType_Input")
losstype.send_keys("Auto")
chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_cboInsuredVehicle_Input").click()
sleep(2)
chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_FinishNavigationTemplateContainerID_btnCreate").click()
sleep(8)
#sleep(5)
#chrome.close()