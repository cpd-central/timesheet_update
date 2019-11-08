from twilio.rest import Client
import datetime
import email_phone_timesheet

year = datetime.datetime.today().year

email_phone_timesheet_dict = email_phone_timesheet.get_dictionary_develop(year)

client = Client("AC3ce1b367bfec1583fb2659f3e24670d7", "d566c629985661e14fc153d897fbe5a2")
twilio_phone_number = "+12029198514"

for user in email_phone_timesheet_dict:
    #get phone number from the dictionary 
    phone_number = email_phone_timesheet_dict[user][1]

    #send text
    try:
        client.messages.create(
            to=phone_number,
            from_=twilio_phone_number,
            body="Hello!  This is a reminder to fill out your timesheet before leaving today.  Thanks!"
        )
    except: 
        print('no phone number!')


