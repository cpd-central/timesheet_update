from pymongo import MongoClient
import xlwings as xw
import pandas as pd
import datetime
import importlib
from dateutil.relativedelta import relativedelta

client = MongoClient('mongodb://192.168.99.110:9999/CEG_PROJECTS_TESTING')
db = client['CEG_PROJECTS_TESTING']

coll = db['timesheets']
coll_users = db['users']

now = datetime.datetime(2019, 9, 11)
now = datetime.datetime.today()
year = now.year
first_payperiod_end = datetime.datetime(2019, 8, 25)

def is_pay_period_due(today):
    time_difference = (today - first_payperiod_end).days 
    if time_difference % 14 == 0:
        return True
    else:
        return False

def get_closest_pay_period(today):
    #checks if it's already a pay period end 
    pay_period_date = is_pay_period_due(today) 
    #loops through until we get a True on the modulus check above 
    while not pay_period_date:
        today = today - datetime.timedelta(days=1)        
        pay_period_date = is_pay_period_due(today)    
    most_recent = today 
    return most_recent 

pay_period_due = is_pay_period_due(now)

if pay_period_due:
    #if we are at the pay period end, we set everyone's sent status to False 
    coll.update_many({'pay_period_sent': {'$exists': True}}, {'$set': {'pay_period_sent': False}}) 

#now, get the data from the database
data = pd.DataFrame(list(coll.find()))

#get the users column from the dataframe
users = data['user']

email_timesheet_dict = {"speichel@ceg-engineers.com": f"C://Users//jjm64//OneDrive//Desktop//timesheet_test_folder//PeichelS.xls",
                        "jmarsnik@ceg-engineers.com": f"C://Users//jjm64//OneDrive//Desktop//timesheet_test_folder//MarsnikJ.xls"}

sheets_dict = {1: "1-January", 2: "2-February", 3: "3-March", 4: "4-April", 5: "5-May", 6:"6-June", 7:"7-July", 8:"8-August",\
              9:"9-September", 10:"10-October", 11:"11-November", 12:"12-December"}

sheets = []
sheets.append(sheets_dict[now.month])  #get current month sheet

#get the date range to use if the timesheet is to be submitted.
most_recent_pay_period_end = get_closest_pay_period(now)
last_pay_period_end = most_recent_pay_period_end - datetime.timedelta(days=14)

if last_pay_period_end.month != now.month:
    sheets.append(sheets_dict[last_pay_period_end.month])

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
            codes = list(data['Codes'][j])
            for code in codes:
                if code == 'Additional Codes':
                    # additional codes is something we use on the laravel end
                    # that we don't need on the python end 
                    codes.remove('Additional_Codes')
                else:
                    # each code can be thought of as a dictionary
                    # with the key being the description and the hours being the
                    # value 
                    keys = data['Codes'][j][code].keys()
                    for key in keys:
                        descriptions.append(key)
             #get the days you will be dealing with for the current sheet. This is needed so it knows how many days in each month to update.
            dates_for_month = []
            if sheet == sheets_dict[now.month]:
                month_day = most_recent_pay_period_end.day
                for i in range(0,14):
                    day = month_day - i
                    if day < 1:
                        break
                    dates_for_month.append(day)
                count_days = len(dates_for_month) 
            else:   
                remaining_days = 14 - count_days 
                print(count_days) 
                ##This gives the last date in the month. 
                end_of_month = last_pay_period_end + relativedelta(day=31)
                for h in range(remaining_days, -1, -1):
                    print(h)         
                    day = end_of_month.day + 1 - h
                    if day > 31:
                        break
                    dates_for_month.append(day)
                    day = day - 1
            print(dates_for_month)           
            expensed_labor = sht.range('A5:A69').value
            code_column = sht.range('AL5:AL69').value
            daterange = sht.range('B3:AF3').value



