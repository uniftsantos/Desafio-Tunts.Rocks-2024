#First of all, you need to set up your environment by enabling your APIs. To do this, you must create a Google Cloud project and
# enable the Google Sheets and Google Drive APIs.
#install the Google client library: pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib

import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError


#The scope for Google refers to the levels of access you will be able to have within the spreadsheet (e.g., read, alter information).
# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# The ID and range of a sample spreadsheet.
#In this variable, we store the spreadsheet ID.
SAMPLE_SPREADSHEET_ID = "1YyTh9UL3b35ExTONbPxySTU-PpJJofgFTsFBPsFxZV8"
#In this variable, we store the range of the spreadsheet.
SAMPLE_RANGE_NAME = "engenharia_de_software!A1:H27"


def main():
 
  creds = None
  #These checks serve to authenticate you with Google. If I haven't done it or if my token is expired, it will redirect me to a page
  # to authenticate my account.  
 
  if os.path.exists("token.json"):
    creds = Credentials.from_authorized_user_file("token.json", SCOPES)
  # If there are no (valid) credentials available, let the user log in   
  if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
      creds.refresh(Request())
    else:
      flow = InstalledAppFlow.from_client_secrets_file(
          "credentials.json", SCOPES
      )
      creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open("token.json", "w") as token:
      token.write(creds.to_json())


  

  try:
    service = build("sheets", "v4", credentials=creds)

    # Call the Sheets API
    #To read my file, I use the get method.
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=SAMPLE_RANGE_NAME).execute()

    #Here we create two empty lists that will later be filled with the final situation and the grade
    values_add_situation  = []
    values_add_grade = []
    
    
    info = result['values']
    #Here I'm going to use a for loop to iterate over the rows of my spreadsheet.
    #I use an if statement to only get the rows with the students.
    for i,rows in enumerate(info):
      grade = 0
      if(i>2):
        #print(rows, i)
        #Here I only get the column of absences.
        absences = int(rows[2])
        #Check to see which students have exceeded 25% of absences.
        if(absences > 60 * 0.25):
            situation = "Reprovados por falta"
            #grade = 0
        else:
            #Average calculation to determine which students passed, took a makeup exam, or repeated the year.
            #the grades range from 0 to 100, if a student has an average lower than 50, they fail due to their grade.
            #If the average is greater than or equal to 50 and less than 70, they have to take a makeup exam. If the average is
            #greater than 70, they pass directly.
            average = (float(rows[3]) + float(rows[4]) + float(rows[5]))/3
            if(average < 50):
                situation ="Reprovado por Nota"
                #grade = 0
            elif(average>=50 and average <70):
                situation = "Exame final"
                grade = 100 - average
                grade = round(grade,0)
            else:
                situation = "Aprovado"
                #grade = 0
        # printing all the students as debug.        
        print(info[i])
        values_add_situation.append([situation])
        values_add_grade.append([grade])
        
       


    #intervalo_linhas = "G{}:G{}".format(min(3), max(27))
    result = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                    range= "G4",valueInputOption="USER_ENTERED",
                    body={'values': values_add_situation}).execute()
    
    result = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                    range= "H4",valueInputOption="USER_ENTERED",
                    body={'values': values_add_grade}).execute()
   
   

    
  except HttpError as err:
    print(err)


if __name__ == "__main__":
  main()