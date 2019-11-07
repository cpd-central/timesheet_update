import pymongo
from pymongo import MongoClient
import datetime
import pandas as pd

now = datetime.datetime.today()
year = now.year

client = MongoClient('mongodb://heroku_0qcgxhh9:f2qrq05120bug3gh44mfqj2ab4@ds131747.mlab.com:31747/heroku_0qcgxhh9')
db = client['heroku_0qcgxhh9']
coll = db['timesheets']

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

#print(code_desc_dict)
#exit()

#updates the reference list in the database.
coll.update_one({
  'name': "reference_list"
},{
  '$set': {
    'codes': code_desc_dict
  }
},upsert=False)




