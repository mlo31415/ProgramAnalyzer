import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request


# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = '1UjHSw-R8dLNFGctUhIQiPr58aAAfBedGznJEN2xBn7o'
SAMPLE_RANGE_NAME = 'A1:Z999'

"""Shows basic usage of the Sheets API.
Prints values from a sample spreadsheet.
"""
creds = None
# The file token.pickle stores the user's access and refresh tokens, and is
# created automatically when the authorization flow completes for the first
# time.
if os.path.exists('token.pickle'):
    with open('token.pickle', 'rb') as token:
        creds = pickle.load(token)
# If there are no (valid) credentials available, let the user log in.
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file(
            'credentials.json', SCOPES)
        creds = flow.run_local_server()
    # Save the credentials for the next run
    with open('token.pickle', 'wb') as token:
        pickle.dump(creds, token)

service = build('sheets', 'v4', credentials=creds)

# Call the Sheets API
sheet = service.spreadsheets()
result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                            range=SAMPLE_RANGE_NAME).execute()
values = result.get('values', [])

if not values:
    raise(ValueError, "No values found")

# The first row is the rooms.
# Make a list of room names and room column indexes
roomIndexes=[]
for i in range(0, len(values[0])):
    if values[0][i] is None:
        break
    if len(values[0][i]) > 0:
        roomIndexes.append(i)

roomNames=values[0]

# Drop the room names from the spreadsheet
values=values[1:]

# Start building the people index
participants={}

# When we find a row with data in column 0, we have found a new time.
rowIndex=0
while rowIndex < len(values):
    row=values[rowIndex]
    if len(row) == 0:   # Skip empty rows
        rowIndex+=1
        continue
    time=row[0]
    for roomIndex in roomIndexes:
        if roomIndex < len(row):
            if len(row[roomIndex]) > 0:
                # OK, this has to be an item name since it's a cell containing text in a row that starts with a timeand in a column that starts with a room
                itemName=row[roomIndex]
                # If there are people scheduled for it, they will be in the next cell down
                if len(values)> rowIndex+1:
                    if len(values[rowIndex+1]) > roomIndex:
                        if len(values[rowIndex+1][roomIndex]) > 0:
                            people=values[rowIndex+1][roomIndex].split(",")
                            for person in people:
                                person=person.strip()
                                if len(person) > 0:
                                    if person not in participants.keys():
                                        participants[person]=[]
                                    participants[person].append((time, roomNames[roomIndex], itemName))
    rowIndex+=2

partlist=participants.keys()
partlist=sorted(partlist, key=lambda x: x.split(" ")[-1:])

txt=open("people with items.txt", "w")
for person in partlist:
    print("", file=txt)
    print(person, file=txt)
    for item in participants[person]:
        print("    "+item[0]+": "+item[2], file=txt)
txt.close()


