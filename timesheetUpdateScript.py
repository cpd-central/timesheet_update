from pymongo import MongoClient
import xlwings as xw
import pandas as pd
import datetime
import importlib
from dateutil.relativedelta import relativedelta


client = MongoClient('mongodb://heroku_0qcgxhh9:f2qrq05120bug3gh44mfqj2ab4@ds131747.mlab.com:31747/heroku_0qcgxhh9')
#client = MongoClient('mongodb://192.168.99.100:9999')

#db = client['heroku_bmf11mmv']
db = client['heroku_0qcgxhh9']

coll = db['timesheets']
coll_users = db['users']

now = datetime.datetime.today()

year = now.year
first_payperiod_end = datetime.datetime(2019, 8, 25)

def is_last_day_of_month(today):
    today_month = today.month
    tomorrow = today + datetime.timedelta(1) 
    tomorrow_month = tomorrow.month
    if tomorrow_month != today_month:
        return True
    else:
        return False

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
is_last_day_of_month = is_last_day_of_month(now)


if pay_period_due:
    #if we are at the pay period end, we set everyone's sent status to False 
    coll.update_many({'pay_period_sent': {'$exists': True}}, {'$set': {'pay_period_sent': False}}) 

#now, get the data from the database
data = pd.DataFrame(list(coll.find()))

#get the users column from the dataframe
users = data['user']

email_timesheet_dict = {
                        "speichel@ceg-engineers.com": f"H://CEG Timesheets//{year}//PeichelS.xls",
                        "jmarsnik@ceg-engineers.com": f"H://CEG Timesheets//{year}//MarsnikJ.xls",
                        "rduncan@ceg-engineers.com": f"H://CEG Timesheets//{year}//DuncanR.xls",
                        "cdolan@ceg.mn": f"H://CEG Timesheets//{year}//DolanC.xls",
                        "kburk@ceg-engineers.com": f"H://CEG Timesheets//{year}//BurkK.xls",
                        "mkaas@ceg-engineers.com": f"H://CEG Timesheets//{year}//KaasM.xls",
                        "bahlsten@ceg-engineers.com": f"H://CEG Timesheets//{year}//AhlstenB.xls",
                        "mbartholomay@ceg-engineers.com": f"H://CEG Timesheets//{year}//BartholomayM.xls",
                        "dborkovic@ceg-engineers.com": f"H://CEG Timesheets//{year}//BorkovicD.xls",
                        "ebryden@ceg-engineers.com": f"H://CEG Timesheets//{year}//BrydenE.xls",
                        "rbuckingham@ceg-engineers.com": f"H://CEG Timesheets//{year}//BuckinghamR.xls",
                        "jcasanova@ceg-engineers.com": f"H://CEG Timesheets//{year}//CasanovaJ.xls",
                        "schowdhary@ceg-engineers.com": f"H://CEG Timesheets//{year}//ChowdharyS.xls",
                        "vince@ceg.mn": f"H://CEG Timesheets//{year}//GranquistV.xls",
                        "nguddeti@ceg-engineers.com": f"H://CEG Timesheets//{year}//GuddetiN.xls",
                        "siqbal@ceg-engineers.com": f"H://CEG Timesheets//{year}//IqbalS.xls",
                        "ajama@ceg-engineers.com": f"H://CEG Timesheets//{year}//JamaA.xls",
                        "skatz@ceg-engineers.com": f"H://CEG Timesheets//{year}//KatzS.xls",
                        "pmalamen@ceg-engineers.com": f"H://CEG Timesheets//{year}//MalamenP.xls",
                        "jmitchell@ceg-engineers.com": f"H://CEG Timesheets//{year}//MitchellJ.xls",
                        "ntmoe@ceg.mn": f"H://CEG Timesheets//{year}//MoeN.xls",
                        "jromero@ceg.mn": f"H://CEG Timesheets//{year}//RomeroJ.xls",
                        #"dsindelar@ceg-engineers.com": f"H://CEG Timesheets//{year}//SindelarD.xls",
                        "turban@ceg-engineers.com": f"H://CEG Timesheets//{year}//UrbanT.xls",
                        "yzhang@ceg-engineers.com": f"H://CEG Timesheets//{year}//ZhangY.xls"
                        }

part_time_users = ["bahlsten@ceg-engineers.com", "schowdhary@ceg-engineers.com", "pmalamen@ceg-engineers.com"] 

sheets_dict = {1: "1-January", 2: "2-February", 3: "3-March", 4: "4-April", 5: "5-May", 6:"6-June", 7:"7-July", 8:"8-August",\
              9:"9-September", 10:"10-October", 11:"11-November", 12:"12-December"}

sheets = []
sheets.append(sheets_dict[now.month])  #get current month sheet

#get the date range to use if the timesheet is to be submitted.
most_recent_pay_period_end = get_closest_pay_period(now)
last_pay_period_end = most_recent_pay_period_end - datetime.timedelta(days=14)

if last_pay_period_end.month != now.month:
    sheets.append(sheets_dict[last_pay_period_end.month])

def write_to_spreadsheet(wb, sheets, month_end):
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
                if code == 'Additional_Codes':
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
                month_day = now.day
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
            ###We create our own index so Xlwings knows where to write the data to.
            #column_index_list: This is for the 1st date being under colomn B, the 2nd being under column C, etc.
            #clear_column_index_list_offset: This is used to wipe rows incase there's no code associated with them and we need the A column.
            column_index_list = ['B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z','AA',\
                            'AB','AC','AD','AE','AF']
            clear_column_index_list_offset = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z','AA',\
                            'AB','AC','AD','AE','AF']

            #Creates an index for the rows on the page.
            row_index_dict = {}
            k = 5
            for exp_lab in expensed_labor:
                if k > 69:
                    break
                if exp_lab == None:
                    k += 1
                    continue
                if exp_lab in row_index_dict:
                    continue
                else:
                    string = str("A"+str(k))
                    row_index_dict.update({sht.range(string).value: k})
                    k += 1
            
            #Nonbillable Codes.
            nonbillable = ["CEG", "CEGTRNG", "CEGEDU", "CEGMKTG"]

            #Because of how the dates are, we need to reformat it to get the day number and insert in the column_index. Subtract 1 to access index 0.
            #@returns the column that date is associated with.
            def get_column_index(day):
                num = day.split('-')[0]
                column = column_index_list[int(num) - 1]
                return column

            #This adds an entry for expensed labor to be billed. These go under Billable Projects.
            def add_expensed_labor(desc,code):
                j = 0
                labor_range = expensed_labor[row_index_dict['Billable Projects: ↓'] - 4:] ##Offset 4 because expensed_labor range is offset A5:A69
                for exp in labor_range:
                    if exp == None:
                        row = row_index_dict['Billable Projects: ↓']+1+j
                        row_index_dict.update( {desc : row} )
                        sht.range(f"A{row}").value = desc
                        sht.range(f"AL{row}").value = code
                        break
                    j += 1

            #This adds an entry for expensed labor not to be billed. This goes under the core 5 nonbillable descriptions.
            def add_nonbillable_labor(desc, code):
                j = 0
                labor_range = expensed_labor[14:19]
                for exp in labor_range:
                    if exp == None:
                        row = 22+j
                        row_index_dict.update( {desc : row} )
                        sht.range(f"A{row}").value = desc
                        sht.range(f"AL{row}").value = code
                        break
                    j += 1

            ##this gets rid of code thats not in the database and also wipes all the columns because they'll be reinserted.
            for m,el in enumerate(expensed_labor):
                #print(row_index_dict) 
                if m < 12 or (m == row_index_dict['Billable Projects: ↓'] - 5): #m == 19    
                    #print(el)
                    continue
                if el in row_index_dict and el in descriptions:
                    row = row_index_dict[el]
                    for z,col in enumerate(clear_column_index_list_offset):
                        if z not in dates_for_month:
                            continue
                        sht.range(f"{col}{row}").value = None
                if el in row_index_dict and el not in descriptions:
                    row = row_index_dict[el]
                    for x,col in enumerate(clear_column_index_list_offset):
                        ###string = str("A"+str(row))
                        if x not in dates_for_month:
                            continue
                        sht.range(f"{col}{row}").value = None
                        #sht.range(f"A{row}").value = None
                        #sht.range(f"AL{row}").value = None


            #This checks if there's any rows with no entries, and if so get rid of them. If the entries are in the database, they get readded later.
            row_range = sht.range('A22:A69').value
            for desc in row_range:
                if desc == None or desc == 'Billable Projects: ↓':
                    continue
                count = 0
                for colm in column_index_list:
                    if sht.range(f"{colm}{row_index_dict[desc]}").value != None:
                        count += 1
                if count == 0:
                    for colm in column_index_list:
                        sht.range(f"{colm}{row_index_dict[desc]}").value = None
                        sht.range(f"A{row_index_dict[desc]}").value = None
                        sht.range(f"AL{row_index_dict[desc]}").value = None
            

            expensed_labor = sht.range('A5:A69').value       #Because we removed some descriptions and codes, lets get an updated list of descriptions.

            #Adds all the code from the database to the timesheet.
            for code in codes:
                dates = list()
                for desc in descriptions:
                    if len(data['Codes'][j][code]) > 0:
                        #print(desc)
                        if desc in data['Codes'][j][code] and len(data['Codes'][j][code][desc]) > 0:
                            dates = data['Codes'][j][code][desc].keys()
                            for day in daterange:
                                if day == None:     #This avoids AF column being called on month that doesn't have a date 31 & Should cover February
                                    continue
                                dt = datetime.datetime.strptime(str(day), '%Y-%m-%d %H:%M:%S')
                                if dt.day not in dates_for_month:     #if date isn't in 2 week period, skip it.
                                    continue
                                day = '{0}-{1}'.format(dt.day, dt.strftime("%b-%y"))
                                for date in dates:
                                    if day == date:
                                        column = get_column_index(day)
                                        if desc not in expensed_labor:
                                                if code not in nonbillable:
                                                    add_expensed_labor(desc, code)
                                                else:
                                                    add_nonbillable_labor(desc, code)                                      

                                                expensed_labor = sht.range('A5:A69').value
                                                code_column = sht.range('AL5:AL69').value
                                        for i,exp in enumerate(expensed_labor):
                                            if exp == desc and code == code_column[i]:
                                                row = row_index_dict[desc]
                                                entry = data['Codes'][j][code][desc][day]
                                                print(f"{date} {desc}: {entry}") 
                                                pay_period_total = pay_period_total + entry 
                                                sht.range(f"{column}{row}").value = entry
                                                if sht.range(f"A{row}").value == None:
                                                    sht.range(f"A{row}").value = desc
                                                    sht.range(f"AL{row}").value = code
                                                break
    return pay_period_total 

def run_macro_and_set_flag(wb, user):
    macro = wb.macro('Sendpayperiodsummary') 
    macro() 
    print('pay period has been sent')
    coll.update_one({'user': user}, {'$set': {'pay_period_sent': True}})
    return None

def check_and_send(wb, pay_period_total, user):
    print('pay period has not been sent')
    print(pay_period_total) 
    if user not in part_time_users: 
        if pay_period_total >= 80: 
            run_macro_and_set_flag(wb, user)
        else:
            print('timesheet needs finishing') 
    else:
        run_macro_and_set_flag(wb, user)
    
    return None 

for j, user in enumerate(users):
    if user not in email_timesheet_dict:
        #if this user does not exist in our dictionary, go to next iteration 
        continue
    #get whether or not the pay period has been sent for this user
    pay_period_sent = data['pay_period_sent'][data['user'] == user].values[0]
    wb = xw.Book(email_timesheet_dict[user])
    app = xw.apps.active

    #we write to the spreadsheet and get the total regardless of the day
    pay_period_total = write_to_spreadsheet(wb, sheets, is_last_day_of_month)
    if pay_period_sent:
        print('pay period sent') 
        pass
    else:
        check_and_send(wb, pay_period_total, user)
    
    wb.save()       #Saves the Spreadsheets.
    print(f"{user} complete")

try:
    app.quit()          #Closes Excel
except NameError:
    print('all timesheets sent')
except:
    print('unknown error!')

