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

dir = os.chdir(r'C:\Users\adam_\Documents\GitHub\Navitas\module-survey')
cnx = sqlite3.connect('moduleSurvey.db')
username = 'a.lowe@sae.edu'

# Delete the file
try:
    os.remove('results-survey683435.xlsx')
except: 
    pass

# Call the web browser
chrome_options = webdriver.ChromeOptions()
prefs = {'download.default_directory' : r'C:\Users\adam_\Documents\GitHub\Navitas\module-survey'} # Changes the download directory
chrome_options.add_experimental_option('prefs', prefs)
chrome_options.add_argument("--window-size=1920, 1080") # Can be changed
chrome_options.add_argument("--headless")
path = 'chromedriver.exe'
browser = webdriver.Chrome(path, options=chrome_options)

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
        browser.find_element_by_css_selector('#user').send_keys(username)
        browser.find_element_by_css_selector('#password').send_keys(keyring.get_password('Module Survey', username))
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
SERVICE_ACCOUNT_FILE = 'keys.json'
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