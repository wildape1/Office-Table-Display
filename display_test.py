import pandas as pd
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow,Flow
from google.auth.transport.requests import Request
import os
import pickle

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


for key, value in engineerNames.iteritems():
    print(value)
    temp_df = active_df[active_df['ENGINEER'] == value]
    temp_df = temp_df.sort_values('JOB #') 
    print(temp_df)
    print()

#print(active_df)
#print(inactive_df)

