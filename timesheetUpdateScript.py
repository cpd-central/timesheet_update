from pymongo import MongoClient
import xlwings as xw
import pandas as pd
import datetime
import importlib
from dateutil.relativedelta import relativedelta
import calendar
import os
import time
import pprint

pp = pprint.PrettyPrinter(indent=2)

client = MongoClient('mongodb://heroku_0qcgxhh9:f2qrq05120bug3gh44mfqj2ab4@ds131747.mlab.com:31747/heroku_0qcgxhh9')
#client = MongoClient('mongodb://localhost:9999')

db = client['heroku_0qcgxhh9']

coll = db['timesheets']
coll_users = db['users']

now = datetime.datetime.today()
#now = datetime.datetime(2019, 12, 15)

year = now.year
month = now.month
day = now.day

first_payperiod_end = datetime.datetime(2019, 8, 25)

def update_reference_list():

    fp = f"H://CEG Timesheets//{year}//CEG Project List.xls"
    sheet = "Open Projects"

    df = pd.read_excel(fp, sheet_name=sheet, header=1)

    columns_to_keep = ['Code', 'Name']
    columns_to_drop = [c for c in df.columns if c not in columns_to_keep]

    df_drop_columns = df.drop(columns=columns_to_drop)
    df_drop_na = df_drop_columns.dropna()

    code_desc_dict = {}
    
    for index, row in df_drop_na.iterrows():
        code_desc_dict[row['Code']] = row['Name']

        #updates the reference list in the database.
        coll.update_one({
            'name': "reference_list"
        }, {
            '$set': {
                'codes': code_desc_dict
            }
        }, upsert=False)
    
    return None

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
                        "speichel@ceg-engineers.com": "PeichelS.xls",
                        "jmarsnik@ceg-engineers.com": "MarsnikJ.xls",
                        "rduncan@ceg-engineers.com": "DuncanR.xls",
                        "cdolan@ceg.mn": "DolanC.xls",
                        "kburk@ceg-engineers.com": "BurkK.xls",
                        "mkaas@ceg-engineers.com": "KaasM.xls",
                        "bahlsten@ceg-engineers.com": "AhlstenB.xls",
                        "mbartholomay@ceg-engineers.com": "BartholomayM.xls",
                        "dborkovic@ceg-engineers.com": "BorkovicD.xls",
                        "ebryden@ceg-engineers.com": "BrydenE.xls",
                        "rbuckingham@ceg-engineers.com": "BuckinghamR.xls",
                        "jcasanova@ceg-engineers.com": "CasanovaJ.xls",
                        "schowdhary@ceg-engineers.com": "ChowdharyS.xls",
                        "vince@ceg.mn": "GranquistV.xls",
                        "nguddeti@ceg-engineers.com": "GuddetiN.xls",
                        "siqbal@ceg-engineers.com": "IqbalS.xls",
                        "ajama@ceg-engineers.com": "JamaA.xls",
                        "skatz@ceg-engineers.com": "KatzS.xls",
                        "pmalamen@ceg-engineers.com": "MalamenP.xls",
                        "jmitchell@ceg-engineers.com": "MitchellJ.xls",
                        "ntmoe@ceg.mn": "MoeN.xls",
                        "jromero@ceg.mn": "RomeroJ.xls",
                        "dsindelar@ceg-engineers.com": "SindelarD.xls",
                        "turban@ceg-engineers.com": "UrbanT.xls",
                        "yzhang@ceg-engineers.com": "ZhangY.xls",
                        "mtuma@ceg-engineers.com": "TumaM.xls"
                       }

part_time_users = ["bahlsten@ceg-engineers.com", "schowdhary@ceg-engineers.com", "pmalamen@ceg-engineers.com"] 

sheets_dict = {1: "1-January", 2: "2-February", 3: "3-March", 4: "4-April", 5: "5-May", 6:"6-June", 7:"7-July", 8:"8-August",\
              9:"9-September", 10:"10-October", 11:"11-November", 12:"12-December"}

sheets = []
sheets.append(sheets_dict[now.month])  #get current month sheet

#get the date range to use if the timesheet is to be submitted.
most_recent_pay_period_end = get_closest_pay_period(now)
last_pay_period_end = most_recent_pay_period_end - datetime.timedelta(days=14)

#nonbillable codes
nonbillable_codes = ['CEG', 'CEGTRNG', 'CEGEDU', 'CEGMKTG']

if last_pay_period_end.month != month:
    sheets.append(sheets_dict[last_pay_period_end.month])

if last_pay_period_end.year != year:
    is_year_crossover = True
else:
    is_year_crossover = False

def get_workbook_and_sheet(sheet_year, user_name, sheet):
    #path = f"H://CEG Timesheets//{sheet_year}//"
    path = f"H://CEG Timesheets//{sheet_year}//"
    full_path = os.path.join(path, user_name)
    print(full_path) 
    wb = xw.Book(full_path)
    app = xw.apps.active
    sht = wb.sheets[sheet]
    return (wb, app, sht)

def write_to_spreadsheet(user_spreadsheet_name, sheets, month_end, user_data, pay_period_sent, is_year_crossover):
    #we also need to map the range of Excel letters to numbers
    letters_to_numbers_dict = {
        0: 'B', 1: 'C', 2: 'D', 3: 'E', 4: 'F',
        5: 'G', 6: 'H', 7: 'I', 8: 'J', 9: 'K',
        10: 'L', 11: 'M', 12: 'N', 13: 'O', 14: 'P',
        15: 'Q', 16: 'R', 17: 'S', 18: 'T', 19: 'U',
        20: 'V', 21: 'W', 22: 'X', 23: 'Y', 24: 'Z',
        25: 'AA', 26: 'AB', 27: 'AC', 28: 'AD', 29: 'AE', 30: 'AF'
    }

    pay_period_total = 0 

    for s, sheet in enumerate(sheets):
        print(sheet) 
        print(is_year_crossover) 
        #if it is a year crossover, check if it's January or December 
        if is_year_crossover:
            #if it's January, set the sheet year to the current year 
            if sheet == '1-January':
                sheet_year = year
            #if it's December, set the sheet year to the previous year 
            if sheet == '12-December':
                sheet_year = last_pay_period_end.year 
        #if it's not a year crossover, set sheet_year to the current year 
        else:
            sheet_year = year

        xw_variables = get_workbook_and_sheet(sheet_year, user_spreadsheet_name, sheet) 
        wb = xw_variables[0]
        app = xw_variables[1]   
        sht = xw_variables[2] 
        
        if sht.range('AF69').value == 'Complete':
            print(f"{sheet} is protected and cannot be written to.")
            wb.save()
            app.quit() 
            continue 
        else:
            #if the sheet isn't protected, we start writing to it            
            #this dictionary will store all of the hours for each project, where the hours dictionary is the value
            # and the key is a tuple of the code and description 
            code_desc_hours_dict = {}
            codes = user_data['Codes'].values[0]
            #we need to get rid of this code, as it is not something CEG uses - it's just for the Laravel system
            try: 
                del codes['Additional_Codes']
            except KeyError:
                print('No additional Codes')
            
            for code in codes:
                keys = list(codes[code].keys()) 
                #each code is a dictionary
                #the key is the description and the value is the hours
                for description in keys:
                    code_desc_tuple = (code, description)
                    code_desc_hours_dict[code_desc_tuple] = codes[code][description] 

        #where everything lives on the spreadsheet        
        date_range = sht.range('B3:AF3').value 

        #remove none values for months with 30 or fewer days
        date_range = [d for d in date_range if d is not None]
        #format our date range into a string that matches what we have in our database 
        date_range_strings = [] 
        for date in date_range:
            date_string = date.strftime("%e-%b-%y").strip(' ') 
            date_range_strings.append(date_string) 

        description_column = 'A'
        code_column = 'AL'
        bill_y_n_column = 'AO'

        ##get the old data and put it into the code_desc_hours_dict before we wipe

        #row_count starts at 17 since that's where the descriptions start
        row_count = 17
        for i in range(row_count, 70):
            code = sht.range(f"AL{i}").value
            description = sht.range(f"A{i}").value
            #if the code exists, we want to see if it has hours 
            if code != None:
                code_desc = (code, description)
                hours = sht.range(f"{letters_to_numbers_dict[date_range[0].day - 1]}{i}:{letters_to_numbers_dict[date_range[-1].day - 1]}{i}").value 
                if all(h is None for h in hours):
                    #if all of the values in the list of hours are none, then there is no time for this code and we can move on 
                    continue 
                #if we get hours, we have to find which days have hours
                else: 
                    #first, check if that code_desc tuple is already in the current data
                    if code_desc not in code_desc_hours_dict:                   
                        code_desc_hours_dict[code_desc] = {} 
            
                    for j, hour in enumerate(hours):
                        if hour != None:
                            date = date_range_strings[j]
                            if date not in code_desc_hours_dict[code_desc]:
                                if isinstance(code_desc_hours_dict[code_desc], list):
                                    code_desc_hours_dict[code_desc] = {}
                                code_desc_hours_dict[code_desc].update({date:hour})

        #wipe everything on sheet

        #wipe descriptions 
        sht.range('A17:A69').value = None
        #wipe time
        #NOTE these are done separately because I think the purple bar after row 15 is causing issues with wiping the whole sheet 
        sht.range('B8:AF15').value = None 
        sht.range('B17:AF69').value = None
        #wipe codes
        sht.range('AL17:AL69').value = None

        #row_count starts at 18 since there's an issue with CEGEDU on line 17  

        row_count = 18 
        #holiday and PTO have special rows
        holiday_row = 14
        pto_row = 15

        for code_desc in code_desc_hours_dict:
            code = code_desc[0]
            description = code_desc[1] 
            # write the code/description in, followed by the time
            sht.range(f"{description_column}{row_count}").value = description 
            sht.range(f"{code_column}{row_count}").value = code              

            #mark whether or not this code is billable in AO
            if code in nonbillable_codes:
                sht.range(f"{bill_y_n_column}{row_count}").value = 'N'
            else:
                sht.range(f"{bill_y_n_column}{row_count}").value = 'Y'

            hours_data = code_desc_hours_dict[code_desc]
        
            #this becomes true if there are any hours in the date range
            # if it stays false, then the row_count won't get incremented and the code/description will get overwritten 
            hours_in_date_range = False
            #loop through each day in the hours for each code
            for hours_day in hours_data:
                #check if the day is in our date range
                #if it's not, we don't need to consider that data on this iteration  
                if hours_day in date_range_strings: 
                    hours_in_date_range = True 
                    hours = hours_data[hours_day]
                    #get the index of this date in our date range list
                    date_index = date_range_strings.index(hours_day)
                    #check if this is holiday or pto hours, as those have special rows 
                    if code_desc == ('CEG', 'Holiday'):
                        sht.range(f"{letters_to_numbers_dict[date_index]}{holiday_row}").value = hours    
                        #decrement row count so we overwrite this row on next iteration
                        if row_count >= 18:
                            #if we're at 17 or higher, we want to bump it down 1 
                            row_count -= 1 
                        else:
                            #if row count is less than 17, we want to put it at 16, so that the row_count += 1 later will bring us back to 17 
                            row_count = 17
                    elif code_desc == ('CEG', 'PTO'):
                        sht.range(f"{letters_to_numbers_dict[date_index]}{pto_row}").value = hours
                        #decrement row count so we overwrite this row on next iteration
                        if row_count >= 18:
                            #if we're at 17 or higher, we want to bump it down 1 
                            row_count -= 1 
                        else:
                            #if row count is less than 17, we want to put it at 16, so that the row_count += 1 later will bring us back to 17 
                            row_count = 17                   
                    else: 
                        sht.range(f"{letters_to_numbers_dict[date_index]}{row_count}").value = hours 
                    #now we match this with our mapping from above to find out which column to put it in in excel

            #increment the row count if there are hours for this code in the date range 
            if hours_in_date_range:
                row_count += 1
        
        #lastly, need to put 'CEG' in  AL17 since it needs to be populated for Donna to send month end
        sht.range(f'{code_column}17').value == 'CEG'
        
        #if we need to check the pay period total, we read from row 70 for the pay period.  
        if not pay_period_sent:
            total_hours_row = 70  
            
            def update_pay_period_total(pay_period_total, start_end):
                current_month_hours = sht.range(f"{letters_to_numbers_dict[start_end[0]]}{total_hours_row}:{letters_to_numbers_dict[start_end[1]]}{total_hours_row}").value 
                print(current_month_hours) 
                #if there's only one day in the month in the pay period, we don't need to sum the hours, just set the total as the single day's hours

                if isinstance(current_month_hours, float):
                    current_month_total = current_month_hours
                else:
                    current_month_total = sum(current_month_hours)
                pay_period_total = pay_period_total + current_month_total
                return pay_period_total 
            
            #check if we have more than one sheet 
            if len(sheets) > 1:
                print('more than one sheet') 
                month_date_range = [] 
                if sheet == sheets_dict[month]:
                    print('current month') 
                    for i in range(1, most_recent_pay_period_end.day + 1):
                        date = datetime.datetime(year, month, i) 
                        date_string = date.strftime("%e-%b-%y").strip(' ')
                        month_date_range.append(date_string)
                    #now, get the indeces of where these dates live in the date range string 
                    #if we're in the current month, the start index is 1 (first of month) 
                    # and the end index is the index of the last day in the month date range   
                    start_end = (0, date_range_strings.index(month_date_range[-1]))  
                    pay_period_total = update_pay_period_total(pay_period_total, start_end) 
                    print(pay_period_total) 
                else:
                    print('previous month') 
                    #if January, we set previous month to 12 (December) 
                    if month == 1:
                        previous_month = 12 
                    else: 
                        previous_month = month - 1 
                    last_day_of_month = calendar.monthrange(last_pay_period_end.year, previous_month)[1]
                    for i in range(last_pay_period_end.day + 1, last_day_of_month + 1): 
                        date = datetime.datetime(last_pay_period_end.year, previous_month, i) 
                        date_string = date.strftime("%e-%b-%y").strip(' ') 
                        month_date_range.append(date_string) 
                    #if we're in the current month, the start index is 1 (first of month) 
                    # and the end index is the index of the last day in the month date range 
                    print(month_date_range) 
                    print(date_range_strings) 
                    start_end = (date_range_strings.index(month_date_range[0]), last_day_of_month - 1) 
                    pay_period_total = update_pay_period_total(pay_period_total, start_end)
                    print(pay_period_total) 
            else:
                print('only one sheet') 
                #if we don't, we can just get the time delta between the most recent pay period end and the last 
                #get the pay period date range
                pay_period_delta = most_recent_pay_period_end - last_pay_period_end
                pay_period_date_range = [] 
                for i in range(pay_period_delta.days):
                    date = last_pay_period_end + datetime.timedelta(days=1) + datetime.timedelta(days=i)
                    date_string = date.strftime("%e-%b-%y").strip(' ')
                    pay_period_date_range.append(date_string) 
                start_end = (date_range_strings.index(pay_period_date_range[0]), date_range_strings.index(pay_period_date_range[-1])) 
                pay_period_total = update_pay_period_total(pay_period_total, start_end)
                print(pay_period_total)

        def find_macro_and_send(wb):
            macro = wb.macro('Sendpayperiodsummary')
            macro()
            return None

        #if we have more than one sheet, we submit on the second sheet 
        if len(sheets) > 1: 
            if s == 1:  
                find_macro_and_send(wb)
        #otherwise, we just submit, since this is the only sheet to submit on 
        else:
            find_macro_and_send(wb)

        wb.save() 
        #give the app some time to fully close 
        app.quit() 
        time.sleep(3) 
    return pay_period_total 

for j, user in enumerate(users): 
    if user not in email_timesheet_dict:
        #if this user does not exist in our dictionary, go to next iteration 
        continue
    print(user)
    #get whether or not the pay period has been sent for this user
    pay_period_sent = data['pay_period_sent'][data['user'] == user].values[0]
    
    user_spreadsheet_name = email_timesheet_dict[user]
 
    #we write to the spreadsheet and get the total regardless of the day
    user_data = data[data['user'] == user]
    pay_period_total = write_to_spreadsheet(user_spreadsheet_name, sheets, is_last_day_of_month, user_data, pay_period_sent, is_year_crossover)
    print(pay_period_total)
    if pay_period_sent:
        print('pay period sent') 
        pass
    else:
        if user not in part_time_users:
            if pay_period_total >= 80:
                coll.update_one({'user': user}, {'$set': {'pay_period_sent': True}})
            else:
                print('timesheet needs finishing')
        else:
            coll.update_one({'user': user}, {'$set': {'pay_period_sent': True}})
    print(f"{user} complete")


#update the reference list
update_reference_list()
