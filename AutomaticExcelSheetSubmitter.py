import xlrd
import time
from selenium.webdriver.common.keys import Keys
import selenium
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import sys

# This program is responsible for automatically inserting data from an excel sheet into a webform,
# and submitting the form for the entire excel sheet. This was done using xlrd, selenium packages.

reload(sys)
sys.setdefaultencoding("utf-8")
#GUI
dir=raw_input('Enter Directory of excel sheet')
z=int(raw_input('Enter start line in excel sheet'))
x=int(raw_input('Enter end line in excel sheet'))

#Open the excel file
workbook = xlrd.open_workbook(dir)
#Open the correct sheet
worksheet = workbook.sheet_by_name('Sheet1')
#OPEN Chrome
driver = webdriver.Chrome()
#OPEN the website
driver.get( "https://go.netotrade.com/80/2584336/2269?t_cmp=(source");
#Loop around all excel sheet data
for i in range(z,x):
    #fetch data
    y = 0
    email = str( worksheet.cell(i,y).value)
    print(email)
    y=1
    name = str( worksheet.cell(i,y).value)
    name = unicode(name, errors='ignore')
    print(name)
    y=2
    country = str(worksheet.cell(i,y).value)
    print(country)
    y=3
    number = str (worksheet.cell(i, y).value).rstrip('.0')
    if number[:6].__eq__("p:+971"):
        number=number[6:]

    print(number)
    #1,0 1,3 then fill , #2,0 2,3 then fill 28 times
    #get form elements

    fullname = driver.find_element_by_name("FirstName");
    mail  =driver.find_element_by_name("Email");
    phone =  driver.find_element_by_name("PhoneNumber");
    submit = driver.find_element_by_css_selector(".submit")
    Intro = driver.find_element_by_name("PhoneCountryCode")
    #fill select item



    #fill other form elements11
    fullname.send_keys(name);
    mail.send_keys(email);
    phone.send_keys(number);
    Intro.send_keys(Keys.CONTROL, "a", Keys.DELETE);
    Intro.send_keys(971)
    #click the submit button
    submit.click();
    #give time to complete request
    time.sleep(5)
    #go back to the previous page
    driver.back()
    #Here,need to identify page elements again in order to clear them
    fullname = driver.find_element_by_name("FirstName");
    mail = driver.find_element_by_name("Email");
    phone = driver.find_element_by_name("PhoneNumber");
    Intro = driver.find_element_by_name("PhoneCountryCode")
    fullname.send_keys(Keys.CONTROL, "a", Keys.DELETE);
    mail.send_keys(Keys.CONTROL, "a", Keys.DELETE);
    phone.send_keys(Keys.CONTROL, "a", Keys.DELETE);
    Intro.send_keys(Keys.CONTROL, "a", Keys.DELETE);




