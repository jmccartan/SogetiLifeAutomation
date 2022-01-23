# python selenium 
# script to enter 40 hours of Sales Support time card
# scheduled job to run Saturday 7:00am 
# using headless chrome - no visible browser - runs against the docker container  
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

browser.get(LifeURL)


delay = 30 # seconds
try:
    myElem = WebDriverWait(browser, delay).until(EC.presence_of_element_located((By.ID, 'unamebean')))
    print ("Page is ready!")
except TimeoutException:
    print ("Loading took too much time!")

try:
    assert "Login" in browser.title
    elem = browser.find_element_by_id("unamebean")
    elem.send_keys(life_user)
    elem = browser.find_element_by_id("pwdbean")
    elem.send_keys(life_pwd)
    elem.send_keys(Keys.RETURN)
    # clicking on links to get to time card 
    players_ele = browser.find_element_by_link_text("Sogeti Time and Expenses").click()
    #print(players_ele)
    search_ele = browser.find_element_by_id("N48").click()
    search_ele = browser.find_element_by_id("Hxccreatetcbutton").click()
    browser.find_element_by_xpath("//select[@name='TemplateList']/option[text()=' -   40 of Sales Support']").click()
    browser.find_element_by_xpath("//button[normalize-space()='Apply Template']").click()
    browser.find_element_by_xpath("//button[normalize-space()='Continue']").click()
    browser.find_element_by_xpath("//button[normalize-space()='Submit']").click()
    emailText = "Time card submission succeeded.  40 hours entered for Sales Support for Week - completed"
    emailSubj = "Completed Weekly Timecard"
except:
    # send fail email
    emailSubj = "Failed Weekly Timecard"
    emailText = "Time card submission failed"

browser.close()

# send email - must be on vpn for this to work 
port = 587  # For starttls
smtp_server = "smtp.office365.com"

sender_email = O365_email_user
receiver_email = O365_email_user
message = MIMEMultipart("alternative")
message["Subject"] = emailSubj
message["From"] = sender_email
message["To"] = receiver_email
# Turn these into plain/html MIMEText objects
part1 = MIMEText(emailText, "html")
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