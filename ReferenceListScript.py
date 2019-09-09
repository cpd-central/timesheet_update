import pymongo
from pymongo import MongoClient
import xlwings as xw
import mongoengine

client = pymongo.MongoClient('mongodb://heroku_bmf11mmv:i6ge501vjrvdv804685mrlhmkf@ds259207.mlab.com:59207/heroku_bmf11mmv')
db = client['heroku_bmf11mmv']
coll = db['timesheets']

wb = xw.Book('H://CEG Timesheets//CEG Projects Master.xls')
app = xw.apps.active
sht = wb.sheets["Project List - Co210"]      
        
code_column = sht.range('A2:A394').value #Needs to be dynamic, not hard coded
desc_column = sht.range('B2:B394').value #Needs to be dynamic, not hard coded

code_desc_dict = {}
for i,code in enumerate(code_column):
    code_desc_dict.update({code: desc_column[i]})
    
entry = {
    "name":"reference_page",
    "codes":code_desc_dict
}

##This is to insert a new reference sheet into the timesheet if one doesn't exist
#coll.insert_one(entry)

#updates the reference list in the database.
coll.update_one({
  'name': "reference_list"
},{
  '$set': {
    'codes': code_desc_dict
  }
},upsert=False)