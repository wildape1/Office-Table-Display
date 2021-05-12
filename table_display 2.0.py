import pandas as pd
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow,Flow
from google.auth.transport.requests import Request
import os
import pickle
import cssFormat

refreshTime = 30
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# here enter the id of your google sheet
SAMPLE_SPREADSHEET_ID_input = '1pcl-Z6OB2btULXQl7RX6UNZmei6vPnuCyDvii5vUua8'
SAMPLE_RANGE_NAME = 'A1:AA1000'

def main():
    global values_input, service
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES) # here enter the name of your downloaded JSON file
            creds = flow.run_local_server(port=0)
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('sheets', 'v4', credentials=creds)

    # Call the Sheets API
    sheet = service.spreadsheets()
    result_input = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID_input,
                                range=SAMPLE_RANGE_NAME).execute()
    values_input = result_input.get('values', [])

    if not values_input and not values_expansion:
        print('No data found.')

main()

df=pd.DataFrame(values_input[1:], columns=values_input[0])
#print(df)

Not_In = df[df['JOB #']=='NOT IN YET'].index.values[0] #finds the index of the not in job section to divide it from Active jobs
#print(Not_In)

active_df = df.iloc[:Not_In,:] #splits the dataframe into 2 
inactive_df = df.iloc[Not_In+1:,:]

active_df = active_df.dropna(subset = ["JOB #"])
active_df = active_df.sort_values('ENGINEER')

engineerNames = active_df["ENGINEER"].drop_duplicates()
#print(engineerNames)

open("active.html","w").close()


active_df_format = pd.DataFrame()
for key, value in engineerNames.iteritems(): 
    temp_df = active_df[active_df['ENGINEER'] == value]
    temp_df = temp_df.sort_values('JOB #')
    active_df_format = pd.concat([active_df_format, temp_df], axis=0) 

    temp_df = temp_df.reset_index()
    #temp_df = temp_df.set_index('JOB #')
    temp_df = temp_df.drop(columns=["index","ENGINEER"])
    temp_df.to_html("temp.html",index=False)

    engineerName = "<h2>Engineer: {} <h2>".format(value)

    with open("temp.html") as temp_file:
      temp_file = temp_file.read()

    with open("active.html") as file:
      file = file.read()

    with open("active.html", "w") as file_to_write:
      file_to_write.write(file + engineerName + temp_file)
    #print(value)
    #print(temp_df)
    #print()


#print(active_df_format)
#print(inactive_df)

with open("active.html") as file:
    file = file.read()

file=file.replace("<table ", "<table class='rwd-table'")
file=file.replace("<th></th> ", "<th style='width:5%;'></th>")
file=file.replace("<th>JOB #</th>", "<th style='width:15%;'>JOB #</th>")
file=file.replace("<th>NAME</th>", "<th style='width:50%;'>NAME</th>")
file=file.replace("<th>PROGRESS DATE</th>", "<th style='width:15%;'>PROGRESS DATE</th>")
file=file.replace("<th>SEAL DATE</th>", "<th style='width:15%;'>SEAL DATE</th>")
file=file.replace("<td>None</td>","<td></td>")

html = """ 
<head>
  <meta http-equiv="refresh" content="{}">
</head> """.format(refreshTime)
with open("active.html", "w") as file_to_write:
    file_to_write.write(cssFormat.html2 + html + file)
    #file_to_write.write(html + file)



#os.startfile("active.html")
