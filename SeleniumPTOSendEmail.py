# python selenium
# extract Saturday next to Saturday last from PTO list and save to htm page
# typically runs Thursday early am 
# docker container jmccartan/life-automation-headless


from secrets import *
from pandas.io import html
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from datetime import date
from datetime import timedelta
from bs4 import BeautifulSoup
import pandas as pd
import os.path
import time
import smtplib
import ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart


LifeURL = "https://life.us.sogeti.com/"

options = webdriver.ChromeOptions()
options.add_argument('headless') #declare it as headless in your code
options.add_argument('no-sandbox') #it run as root
browser = webdriver.Chrome(chrome_options=options)

# browser = webdriver.Chrome()
browser.get(LifeURL)

delay = 3  # seconds
try:
    myElem = WebDriverWait(browser, delay).until(
        EC.presence_of_element_located((By.ID, 'unamebean')))
    print("Page is ready!")
except TimeoutException:
    print("Loading took too much time!")


#delay = 10
assert "Login" in browser.title
elem = browser.find_element_by_id("unamebean")
elem.send_keys(life_user)
elem = browser.find_element_by_id("pwdbean")
elem.send_keys(life_pwd)
elem.send_keys(Keys.RETURN)
# clicking on links to get to time card
players_ele = browser.find_element_by_link_text(
    "Sogeti Employee - Self Service").click()
players_ele = browser.find_element_by_link_text(
    "Sogeti Leave Request Approval").click()
# set dates
today = date.today()
# assumption is runs on Thursday - adding 2 gets to Saturday
Saturday = today + timedelta(days=2)
SaturdayNextFormatted = Saturday.strftime("%d-%b-%Y")
LastSaturday = Saturday + timedelta(days=-7)
SaturdayLastFormatted = LastSaturday.strftime("%d-%b-%Y")

# temp reassign of date parms - for Kate 4th of July week
#SaturdayLastFormatted = "25-Dec-2021"
#SaturdayNextFormatted = "02-Jan-2022"


elem = browser.find_element_by_id("searchLeaveDateRangeStart")
elem.send_keys(SaturdayLastFormatted)
elem = browser.find_element_by_id("searchLeaveDateRangeEnd")
elem.send_keys(SaturdayNextFormatted)
browser.find_element_by_id("SearchBtn").click()
delay = 5
page_source_overview = browser.page_source
soup = BeautifulSoup(page_source_overview, 'lxml')
pto_table = soup.find("table", attrs={"class": "x1o"})
pto_html_text = str(pto_table)
pto_header = "PTO from " + SaturdayLastFormatted + " to " + SaturdayNextFormatted

full_detail_body = pto_html_text
# print to console
print(pto_html_text)

df = pd.read_html(str(pto_html_text))[0] 

# remove columns - starting at 0 - count
df.drop(df.columns[[0, 6, 8, 9, 10, 11, 12, 13, 14, 15]],axis=1, inplace=True)


pto_html_text = df.to_html(classes = 'table table-striped')
#pto_html_text = df 


# send email
port = 587  # For starttls
smtp_server = "smtp.office365.com"


sender_email = O365_email_user
receiver_email = PTO_dist_list
message = MIMEMultipart("alternative")
message["Subject"] = pto_header
message["From"] = sender_email
message["To"] = receiver_email
# Turn these into plain/html MIMEText objects
part1 = MIMEText(pto_html_text, "html")
# Add HTML/plain-text parts to MIMEMultipart message
# The email client will try to render the last part first
message.attach(part1)

print("getting ready to send")
context = ssl.create_default_context()
with smtplib.SMTP(smtp_server, port) as server:
    server.ehlo()  # Can be omitted
    server.starttls(context=context)
    server.ehlo()  # Can be omitted
    server.login(O365_email_user, O365_email_pwd)
    server.send_message(message)
    #server.sendmail(sender_email, receiver_email, message)

browser.close()

# sending email with full details to me only 
message2 = MIMEMultipart("alternative")
message2["Subject"] = pto_header
message2["From"] = sender_email
message2["To"] = PTO_dist_list_just_me
part2 = MIMEText(full_detail_body, "html")
message2.attach(part2)

with smtplib.SMTP(smtp_server, port) as server2:
    server2.ehlo()  # Can be omitted
    server2.starttls(context=context)
    server2.ehlo()  # Can be omitted
    server2.login(O365_email_user, O365_email_pwd)
    server2.send_message(message2)
