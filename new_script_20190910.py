from pymongo import MongoClient
import xlwings as xw
import pandas as pd
import datetime
import importlib
from dateutil.relativedelta import relativedelta

client = MongoClient('mongodb://192.168.99.100:9999/heroku_bmf11mmv')
db = client['heroku_bmf11mmv']

coll = db['timesheets']
coll_users = db['users']

now = datetime.datetime(2019, 9, 7)
year = now.year

first_payperiod_end = datetime.datetime(2019, 8, 24)
time_difference = now - first_payperiod_end

if time_difference.days % 14 == 0:
    pay_period_due = True
else:
    pay_period_due = False

if pay_period_due:
    #if we are at the pay period end, we set everyone's sent status to False 
    coll.update_many({'pay_period_sent': {'$exists': True}}, {'$set': {'pay_period_sent': False}}) 

#now, get the data from the database
data = pd.DataFrame(list(coll.find()))
#get the users column from the dataframe
users = data['user']

email_timesheet_dict = {"speichel@ceg-engineers.com": f"C://Users//jmarsnik//Desktop//timesheet_test_folder//PeichelS.xls",
                        "jmarsnik@ceg-engineers.com": f"C://Users//jmarsnik//Desktop//timesheet_test_folder//MarsnikJ.xls"}

sheets_dict = {1: "1-January", 2: "2-February", 3: "3-March", 4: "4-April", 5: "5-May", 6:"6-June", 7:"7-July", 8:"8-August",\
              9:"9-September", 10:"10-October", 11:"11-November", 12:"12-December"}

sheets = []
sheets.append(sheets_dict[now.month])  #get current month sheet

#get the date range to use if the timesheet is to be submitted.
date_previous = now - datetime.timedelta(days=14)
if date_previous.month != now.month:
    sheets.append(sheets_dict[date_previous.month])

for j, user in enumerate(users):
    if user not in email_timesheet_dict:
        #if this user does not exist in our dictionary, go to next iteration 
        continue
    #get whether or not the pay period has been sent for this user
    pay_period_sent = data['pay_period_sent'][data['user'] == user].values[0]
    if pay_period_sent:
        #if the pay period has been sent, we go to the next user
        continue
    
    #if not, we open the workbook and enter/submit
    wb = xw.Book(email_timesheet_dict[user])
    app = xw.apps.active

    #set the pay period total to 0
    pay_period_total = 0
    for sheet in sheets:
        sht = wb.sheets[sheet] 
        if sht.range('AF69').value == "Complete":
            print(f"{sheet} is prtected and can't be written to.")
        else:
            descriptions = []
            codes = data['Codes'][data['user'] == user] 
            print(codes) 
            codes = list(data['Codes'][j])
            print(codes)
