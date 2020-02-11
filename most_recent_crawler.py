import pandas as pd
import pymongo
import xlwings as xw
import time
from datetime import datetime
import calendar

#in case the spreadsheet is missing projects
#between the master and the spreadsheet itself, we should have everything
master_fp = r'Z:\CEG\Timesheets\CEG Timesheets\CEG Projects Master.xls'
master_df = pd.read_excel(master_fp)
master_ser = master_df['ACTIVITY']

#def merge_hours_to_projects():

def get_newest_hours():
    today = datetime.today()
    month_number = today.month 
    months = list()
    for x in range(0, month_number):
        month = calendar.month_name[x + 1]
        months.append(month)

    year = str(today.year)

    fp = f'Z://CEG//Timesheets//CEG Timesheets//{year}//{year} Hours by Project.xlsx'

    projects_df = pd.read_excel(fp, sheet_name='YTD by Person by Project', header=2)
    projects_ser = projects_df['Project'].dropna()
    all_projects_ser = projects_ser.append(master_ser).drop_duplicates()
    all_projects_ser = all_projects_ser[all_projects_ser != '** some projects missing']
    print(all_projects_ser) 
    #project_codes = all_projects_ser.sort_values().tolist()	
    project_codes = all_projects_ser.tolist()

    wb = xw.Book(fp)
    sht = wb.sheets['Project Report']
    app = xw.apps.active
    
    for i, code in enumerate(project_codes):
        sht.range('E2').value = code

        code_check = sht.range('E2').value
        print(code_check)

        df = sht.range('A6:CA19').options(pd.DataFrame).value

        hours_df = df.iloc[:12, :]
        hours_df.dropna(inplace=True, axis='columns')
        hours_df.columns = hours_df.columns.fillna('noname')

        for month in hours_df['Month']:
            month = month.split('-')[1]

        hours_df['Month'] = hours_df['Month'].str.split('-', expand=True)[1]
        hours_df.set_index('Month', inplace=True)
        #print(hours_df)
        existing_months_df = hours_df.loc[months, :]
        #print(existing_months_df)
        
        existing_months_df_transpose = existing_months_df.transpose()
        existing_months_dict = existing_months_df_transpose.to_dict()

        host = 'ds131747.mlab.com'
        #host = 'localhost'
        port = '31747'
        #port= '9999'
        user = 'heroku_0qcgxhh9'
        password = 'f2qrq05120bug3gh44mfqj2ab4'
        db_name = 'heroku_0qcgxhh9'
        #db_name = 'CEG_PROJECTS_TESTING'

        client = pymongo.MongoClient(f'mongodb://{user}:{password}@{host}:{port}/{db_name}')
        #client = pymongo.MongoClient(f'mongodb://{host}:{port}')
        db = client[db_name]
        coll = db['projects']
        all_hours_coll = db['hours_by_project']

        coll.update_one({'projectcode': code}, {'$set': {f'hours_data.{year}': existing_months_dict}})

        all_hours_coll.update_one({'code': code}, {'$set': {f'hours_data.{year}': existing_months_dict}}, upsert=True)
    wb.save() 
    app.quit() 
    #coll.insert_one({'hello': 'world'})
get_newest_hours()



