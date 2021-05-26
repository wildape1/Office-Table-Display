import pandas as pd
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow,Flow
from google.auth.transport.requests import Request
import os, time
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

def updatefile():
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

  open("activetemp.html","w").close()


  active_df_format = pd.DataFrame()
  for key, value in engineerNames.iteritems(): 
      temp_df = active_df[active_df['ENGINEER'] == value]
      temp_df = temp_df.sort_values('JOB #')
      active_df_format = pd.concat([active_df_format, temp_df], axis=0) 

      temp_df = temp_df.reset_index()
      #temp_df = temp_df.set_index('JOB #')
      temp_df = temp_df.drop(columns=["index","ENGINEER"])
      temp_df.to_html("temp.html",index=False)

      engineerName = "<h2>- {} -<h2>".format(value)

      with open("temp.html") as temp_file:
        temp_file = temp_file.read()

      with open("activetemp.html") as file:
        file = file.read()

      with open("activetemp.html", "w") as file_to_write:
        file_to_write.write(file + engineerName + temp_file)
      #print(value)
      #print(temp_df)
      #print()

  inactive_df = inactive_df.drop(columns=["ENGINEER","PROGRESS 1","PROGRESS 2","SEALED"])
  inactive_df.to_html("inactive.html",index=False)

  with open("inactive.html") as inactive_file:
    inactive_file = inactive_file.read()

  with open("activetemp.html") as file:
    file = file.read()

  html = "<h2> NOT IN YET <h2>"

  with open("activetemp.html", "w") as file_to_write:
    file_to_write.write(file + html + inactive_file)

  #print(active_df_format)
  #print(inactive_df)

  with open("activetemp.html") as file:
      file = file.read()

  file=file.replace("<table ", "<table class='rwd-table'")
  file=file.replace("<th></th> ", "<th style='width:1%;'></th>")
  file=file.replace("<th>JOB #</th>", "<th style='width:12%;'>JOB #</th>")
  file=file.replace("<th>NAME</th>", "<th style='width:60%;'>NAME</th>")
  file=file.replace("<th>PROGRESS 1</th>", "<th style='width:12%;'>PROG 1</th>")
  file=file.replace("<th>PROGRESS 2</th>", "<th style='width:12%;'>PROG 2</th>")
  file=file.replace("<th>SEALED</th>", "<th style='width:10%;'>SEALED</th>")
  file=file.replace("<td>None</td>","<td></td>")

  refreshHtml = """ 
  <head>
    <meta http-equiv="refresh" content="{}">
  </head> """.format(refreshTime)
  with open("activetemp.html", "w") as file_to_write:
      file_to_write.write(cssFormat.html2 + file)
      #file_to_write.write(html + file)

  with open("activetemp.html") as file:
      file = file.read()

  with open("active.html", "w") as file_2_write:
      file_2_write.write(file + refreshHtml)



#os.startfile("active.html")


while(1):
  main()
  updatefile()
  #print("update")
  time.sleep(30)
