from selenium import webdriver
import time
import pandas as pd
from googleapiclient.discovery import build
from google.oauth2 import service_account
from tqdm import trange
import sys
import os
from progress.bar import Bar
import sqlite3
import keyring
from datetime import datetime
import logevent

os.chdir(os.getcwd())
dir = os.getcwd()

# Delete the file
try:
    os.remove('results-survey683435.xlsx')
except: 
    pass

# Variables
modRegDB = os.environ.get('modRegDB') # Access the system Enviroment Variables to get the modReg SQL path
attendanceDB = os.environ.get('attendanceDB')
chromedriver = os.environ.get('chromedriverPath')
NAV_USER = os.environ.get('NAV_USER')
NAV_PASS = os.environ.get('NAV_PASS')
SURVEY_USER = os.environ.get('SURVEY_USER')
SURVEY_PASS= os.environ.get('SURVEY_PASS')
emailAddress = os.environ.get('EMAIL_USER')
emailPassword = os.environ.get('EMAIL_PASS')
keysJSON = os.environ.get('keysJSON')
moduleSurveyDB = os.environ.get('moduleSurveyDB')

# Connect to the SQL database
cnx = sqlite3.connect(moduleSurveyDB)

# Call the web browser
chrome_options = webdriver.ChromeOptions()
prefs = {'download.default_directory' : str(dir)} # Changes the download directory
chrome_options.add_experimental_option('prefs', prefs)
chrome_options.add_argument("--window-size=1920, 1080") # Can be changed
chrome_options.add_argument("--headless")
browser = webdriver.Chrome(chromedriver, options=chrome_options)

# Navigate to the chosen website 
try:
    browser.get('https://survey.sae.edu/index.php/admin/export/sa/exportresults/surveyid/683435')
    time.sleep(2)
except Exception as error:
    logevent.logEvent.failLog(error)
    browser.quit()
    sys.exit()
   
# This downloads the data
bar = Bar('Downloading file...', max=30)
def exportdata():
    try:
        browser.find_element_by_css_selector('#user').send_keys(SURVEY_USER)
        browser.find_element_by_css_selector('#password').send_keys(SURVEY_PASS)
        browser.find_element_by_xpath('//*[@id="loginform"]/div[2]/div/p/button').click()
        browser.find_element_by_css_selector('#xls').click()
        browser.find_element_by_css_selector('#panel-4 > div.panel-body > div:nth-child(1) > div > label:nth-child(4)').click()
        submit = browser.find_element_by_css_selector('#export-button')
        browser.minimize_window()
        submit.click()
    except Exception as error:
        logevent.logEvent.failLog(error)
        browser.quit()
        sys.exit()
        
# Function executes
exportdata()

# Bar visual
for i in range(30):
    time.sleep(1)
    bar.next()

browser.quit()

# Read & filter the downloaded file
print("Filtering...")
time.sleep(2) # Wait 2 seconds
try:
    df = pd.read_excel('results-survey683435.xlsx')
    df = df[df['TP. Teaching Period'].isin(["20T3","21T1","21T2"]) & df['Campus. Campus'].isin(["Liverpool", "London", "Oxford", "Glasgow","Online"])].dropna(how='all')
except Exception as error:
    logevent.logEvent.failLog(error)
    sys.exit()

# Google Sheets Setup
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SERVICE_ACCOUNT_FILE = keysJSON
credentials = None
credentials = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)

# The ID and range of a sample spreadsheet.
modulesurveydata = '1UCYodN9q1MYt3embI-oelo9994ywSpYolnkGduzA4JM'
service = build('sheets', 'v4', credentials=credentials)

# Call the Sheets API and write data
df.fillna('', inplace=True)
sheet = service.spreadsheets()
data = df.values.tolist()
def updatedata():
    print("Uploading data...")
    sheet.values().update(spreadsheetId=modulesurveydata, range="survey data!A2", valueInputOption="USER_ENTERED", body={"values":data}).execute()
    df.to_sql('moduleSurvey', con=cnx,if_exists='replace')
    print("Done.")

# Function executes
updatedata()

# Delete the file
os.remove('results-survey683435.xlsx')

print("Process complete.")
logevent.logEvent.successLog()
sys.exit()