# Project Dashboard 
# this one scheduled on windows machine for saturday morning 7am
# docker run --rm -v /volume1/WeeklyPythonScripts:/seleniumdata mccartan/python-selenium-headless-pandas python3 -u ./[name of python script].py

from secrets import *
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from datetime import date
import time 
import smtplib
import ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.utils import formatdate
from email import encoders
from bs4 import BeautifulSoup
import pandas as pd 
import numpy as np 
import openpyxl
import xlsxwriter


LifeURL = "https://life.us.sogeti.com/"

options = webdriver.ChromeOptions()
options.add_argument('headless') 
options.add_argument('no-sandbox') 
browser = webdriver.Chrome(chrome_options=options)

browser.get(LifeURL)

# life sometimes takes a bit to load - have a delay to wait for element on page 
delay = 30 # seconds
try:
    myElem = WebDriverWait(browser, delay).until(EC.presence_of_element_located((By.ID, 'unamebean')))
    print ("Page is ready!")
except TimeoutException:
    print ("Loading took too much time!")

# try:
assert "Login" in browser.title
elem = browser.find_element_by_id("unamebean")
elem.send_keys(life_user)
elem = browser.find_element_by_id("pwdbean")
elem.send_keys(life_pwd)
elem.send_keys(Keys.RETURN)

players_ele = browser.find_element_by_link_text("Sogeti Employee - Self Service").click()

search_ele = browser.find_element_by_id("N77").click()

search_ele = browser.find_element_by_id("SearchEndFromDate")
#today = date.today()
ProjEndDate = date.today().strftime("%d-%b-%Y")
search_ele.send_keys(ProjEndDate)
browser.find_element_by_xpath("//select[@id='SearchOrg']/option[text()='S-Iowa']").click()

browser.find_element_by_xpath("//*[@id='CustomSimpleSearch']/tbody/tr[2]/td[2]/table/tbody/tr[2]/td[3]/button[1]").click()

page_source_overview = browser.page_source
soup = BeautifulSoup(page_source_overview, 'lxml')
pdb_table = soup.find("table", attrs={"class": "x1o"})
pdb_table_html = str(pdb_table)

tableIA = pdb_table_html

table_iowa = pd.read_html(str(pdb_table_html))
df_IA = table_iowa[0]
cols = [("KPI's", 'As-Sold Margin'), ("KPI's", 'Re-Baselined Margin'), ("KPI's", 'EAC Revenue'), ("KPI's", 'EAC Cost'), ("KPI's", 'EAC DVI'), ('Financial Details', 'Sales Amount'), ('Financial Details', 'Revenue to Date'), ('Financial Details', 'Cost to Date'), ('Financial Details', 'Revenue at Completion'), ('Financial Details', 'Cost at Completion')]
#df_IA[cols] = df_IA[cols].apply(pd.to_numeric, errors='coerce')
df_IA[cols] = df_IA[cols].replace({'\$': '', ',': ''}, regex=True).astype(float)

browser.find_element_by_xpath("//select[@id='SearchOrg']/option[text()='S-Nebraska']").click()
browser.find_element_by_xpath("//*[@id='CustomSimpleSearch']/tbody/tr[2]/td[2]/table/tbody/tr[2]/td[3]/button[1]").click()


page_source_overview = browser.page_source
soup = BeautifulSoup(page_source_overview, 'lxml')
pdb_table = soup.find("table", attrs={"class": "x1o"})
pdb_table_html = str(pdb_table)

tableNE = pdb_table_html
table_nebraska = pd.read_html(str(pdb_table_html))
df_NE = table_nebraska[0]
df_NE[cols] = df_NE[cols].replace({'\$': '', ',': ''}, regex=True).astype(float)


ExcelFileName = "IA NE Project Dashboard "  + str(ProjEndDate) + ".xlsx"
IASheetname = "Iowa " + str(ProjEndDate)
NESheetname = "Nebraska " + str(ProjEndDate)

DashboardWriter = pd.ExcelWriter(ExcelFileName, engine="xlsxwriter")
df_IA.to_excel(DashboardWriter, sheet_name = IASheetname)
df_NE.to_excel(DashboardWriter, sheet_name = NESheetname)

thisWorkbook = DashboardWriter.book
IASheet = DashboardWriter.sheets[IASheetname]
NESheet = DashboardWriter.sheets[NESheetname]


# Adding simple number format - keeping in as an example 
#fmt_number = thisWorkbook.add_format({ “num_format” : “0” })
# Adding currency format
fmt_currency = thisWorkbook.add_format({ 'num_format' : '$#,##0.00' ,'bold' :False })
fmt_center = thisWorkbook.add_format({ 'align' : 'center'})


# yes this should be done by iterating over a collection rather than repeating :) 
IASheet.set_column("N:O", 10, fmt_currency)
NESheet.set_column("N:O", 10, fmt_currency)

IASheet.set_column("Q:T", 10, fmt_currency)
NESheet.set_column("Q:T", 10, fmt_currency)

IASheet.set_column("V:W", 10, fmt_currency)
NESheet.set_column("V:W", 10, fmt_currency)

IASheet.set_column("V:W", 10, fmt_currency)
NESheet.set_column("V:W", 10, fmt_currency)

IASheet.set_column("P:P", 10, fmt_center)
NESheet.set_column("P:P", 10, fmt_center)

IASheet.set_column("U:U", 10, fmt_center)
NESheet.set_column("U:U", 10, fmt_center)

IASheet.set_column("X:X", 10, fmt_center)
NESheet.set_column("X:X", 10, fmt_center)

#clear garbage from screen scrape 
IASheet.write('B1', ' ')
NESheet.write('B1', ' ')

DashboardWriter.save()
print('xls saved... starting email')

emailText = tableIA + '\n\n\n' + tableNE

browser.close()

emailSubj = "IA NE Project Dashboard " + str(ProjEndDate)

port = 587  # For starttls
smtp_server = "smtp.office365.com"

sender_email = O365_email_user
receiver_email = ProjectDashboardDistList
message = MIMEMultipart("alternative")
message["Subject"] = emailSubj
message["From"] = sender_email
message["To"] = receiver_email
# Turn these into plain/html MIMEText objects

part1 = MIMEText(emailText, "html")
# Add HTML/plain-text parts to MIMEMultipart message
# The email client will try to render the last part first
message.attach(part1)
print("adding attachment part2")
part2 = MIMEBase('application', "octet-stream")
part2.set_payload(open(ExcelFileName, "rb").read())
encoders.encode_base64(part2)
part2.add_header('Content-Disposition', 'attachment; filename="' + ExcelFileName + '"')
message.attach(part2)
print('after trying to attach xlx')
print("getting ready to send")
context = ssl.create_default_context()
with smtplib.SMTP(smtp_server, port) as server:
    server.ehlo()  # Can be omitted
    server.starttls(context=context)
    server.ehlo()  # Can be omitted
    server.login(O365_email_user, O365_email_pwd)
    server.send_message(message)