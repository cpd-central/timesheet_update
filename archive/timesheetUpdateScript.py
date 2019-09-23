import pymongo 
from pymongo import MongoClient 
import xlwings as xw 
import pandas as pd 
import datetime 
import importlib 
from dateutil.relativedelta import relativedelta 

client = pymongo.MongoClient('mongodb://heroku_bmf11mmv:i6ge501vjrvdv804685mrlhmkf@ds259207.mlab.com:59207/heroku_bmf11mmv')
db = client['heroku_bmf11mmv']
coll = db['timesheets']
coll_users = db['users']

data = pd.DataFrame(list(coll.find()))
users = data['user']



now = datetime.datetime.now()
year = now.year

#have the payperiod submit be sunday morning.
first_payperiod_end = datetime.datetime(2019, 8, 25)

time_difference = now - first_payperiod_end

if time_difference.days % 14 == 0:
    pay_period_due = True
else:
    pay_period_due = False



email_timesheet_dict = {"speichel@ceg-engineers.com": f"H://CEG Timesheets//{year}//PeichelS.xls",
                        "jmarsnik@ceg-engineers.com": f"H://CEG Timesheets//{year}//MarsnikJ.xls"}
#email_timesheet_dict = {"jmarsnik@ceg-engineers.com": f"H://CEG Timesheets//{year}//MarsnikJ - Copy.xls"}



sheets_dict = {1: "1-January", 2: "2-February", 3: "3-March", 4: "4-April", 5: "5-May", 6:"6-June", 7:"7-July", 8:"8-August",\
              9:"9-September", 10:"10-October", 11:"11-November", 12:"12-December"}

sheets = []
sheets.append(sheets_dict[now.month])  #get current month sheet

##If the two week time period overlaps two months, get the month previous and append it to the Sheets array to use in the For Loop.##
date_previous = datetime.datetime.now() - datetime.timedelta(days=14)      #gets 14 days prior to now
if date_previous.month != now.month:                #if the months aren't the same, get the sheet of the pervious month
    sheets.append(sheets_dict[date_previous.month])

count_days = 0
for j,user in enumerate(users):
    if user not in email_timesheet_dict:
        continue
    wb = xw.Book(email_timesheet_dict[user])
    app = xw.apps.active
    
    two_week_total = 0 
    for sheet in sheets:
        sht = wb.sheets[sheet]
        if sht.range('AF69').value == "Complete":
            print(f'{sheet} is protected and can''t be written to.')
        else:
            descriptions = []
            codes = list(data['Codes'][j])
            for code in codes:
                if code == 'Additional_Codes':
                    codes.remove('Additional_Codes')        ##removes Additional_Codes cause it's not  a code
                else:
                    keys = data['Codes'][j][code].keys()    ##Gets the description under the Key
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
                    count_days += 1
                    dates_for_month.append(day)
            else:
                count_days = 14 - count_days 
                for h in range(count_days + 1, -1, -1):
                    date = date_previous + relativedelta(day=31)        ##This gives the last date in the month.
                    day = date.day - h
                    if day > 31:
                        break
                    dates_for_month.append(day)
                    day = day + 1

            expensed_labor = sht.range('A5:A69').value
            code_column = sht.range('AL5:AL69').value
            daterange = sht.range('B3:AF3').value

            ##We create our own index so Xlwings knows where to write the data to.
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
                j = 0;
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
                j = 0;
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
                if m < 12 or (m == row_index_dict['Billable Projects: ↓'] - 5): #m == 19    
                    print(el)
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
                        print(desc)
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
                                                two_week_total = two_week_total + entry 
                                                sht.range(f"{column}{row}").value = entry
                                                if sht.range(f"A{row}").value == None:
                                                    sht.range(f"A{row}").value = desc
                                                    sht.range(f"AL{row}").value = code
                                                break
                print(f"{sheet} {code} complete")
    
    #run the submit pay period if it's two weeks from the first pay period date 
        if pay_period_due: 
            if two_week_total >= 80: 
                macro = wb.macro('Sendpayperiodsummary') 
                macro() 
        else:
            print('timesheet needs finishing')
    wb.save()       #Saves the Spreadsheets.
    print(f"{user} complete")


app.quit()          #Closes Excel
