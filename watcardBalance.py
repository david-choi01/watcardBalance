from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from datetime import date
from datetime import timedelta
import re, openpyxl

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

def startDate():
    today = date.today()
    startDate = today - timedelta(days=6)
    return startDate.strftime('%m/%d/%Y')

def transactionSource():
    url = "https://watcard.uwaterloo.ca/OneWeb/Account/LogOn"
    account = ENTER_ACCOUNT
    pin = ENTER_PASSWORD
    # TODO: Login into WatCard Balance Website
    browser = webdriver.Firefox(executable_path=r'PATH TO GECKODRIVER')
    browser.implicitly_wait(20)
    browser.get(url)
    accountElem = browser.find_element_by_id("Account")
    accountElem.send_keys(account)
    pinElem = browser.find_element_by_id("Password")
    pinElem.send_keys(pin)
    pinElem.submit()
    # TODO: Open transaction history
    transactionElem = browser.find_element_by_xpath("//a[@href='../Financial/Transactions']")
    transactionElem.click()
    dateStartElem = browser.find_element_by_id("trans_start_date")
    dateStartElem.clear()
    dateStartElem.send_keys(startDate())
    searchElem = browser.find_element_by_id("trans_search")
    searchElem.click()
    # TODO: Download page source
    pageSource = browser.page_source
    browser.quit()
    return pageSource

def dataScrape(source):
    transactionData = {}
    today = date.today()
    rangeStart = 0
    for i in range(0,6):
        currentDate = today - timedelta(days=i)
        dateRegex = re.compile(r'(<.*>)(' + currentDate.strftime('%m/%d/%Y') + ')(.*</td>)')
        if dateRegex.findall(source):
            date_temp = dateRegex.findall(source)
            transactionNumber = len(date_temp)
            rangeEnd = rangeStart + transactionNumber
            transRegex = re.compile(r'(?:-)?\$(?:\d)+\.(?:\d)+')
            transaction_temp = transRegex.findall(source)
            transactionData.update({currentDate.strftime('%m/%d/%Y'): transaction_temp[rangeStart:rangeEnd]})
            rangeStart = rangeEnd
        else:
            continue
    return transactionData        

def dataSave(transactionData):
    wb = openpyxl.Workbook()
    sheet = wb.get_active_sheet()
    cellIndex = 1
    for k,v in transactionData.items():
        for i in range(len(v)):
            sheet.cell(row=cellIndex, column=1).value = k
            sheet.cell(row=cellIndex, column=2).value = v[i]
            cellIndex = cellIndex + 1
    wb.save("watcardTransaction.xlsx")

def dataSend():
    emailFrom = EMAIL
    emailTo = EMAIL

    message = MIMEMultipart()
    message['From'] = emailFrom
    message['To'] = emailTo
    message['Subject'] = "Watcard Balance Data"

    fileName = "watcardTransaction.xlsx"
    emailAttachment = open("PATH TO ATTACHMENT", "rb")
    part = MIMEBase('application', 'octet-stream')
    part.set_payload((emailAttachment).read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', 'attachment; filename = %s' % fileName)
    message.attach(part)
    
    
    smtpObj = smtplib.SMTP('smtp.gmail.com', 587)
    smtpObj.ehlo()
    smtpObj.starttls()
    smtpObj.login(EMAIL, PASSWORD)
    smtpObj.sendmail(emailFrom, emailTo, message.as_string())
    smtpObj.quit()
    
        

source = transactionSource()
transactionData = dataScrape(source)
dataSave(transactionData)
dataSend()
