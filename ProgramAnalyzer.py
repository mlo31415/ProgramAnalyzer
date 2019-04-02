import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request


# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

# The ID and range of a sample spreadsheet.
SPREADSHEET_ID ='1UjHSw-R8dLNFGctUhIQiPr58aAAfBedGznJEN2xBn7o'

creds = None
# The file token.pickle stores the user's access and refresh tokens, and is
# created automatically when the authorization flow completes for the first time.
if os.path.exists('token.pickle'):
    with open('token.pickle', 'rb') as token:
        creds = pickle.load(token)

# If there are no (valid) credentials available, let the user log in.
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
        creds = flow.run_local_server()
    # Save the credentials for the next run
    with open('token.pickle', 'wb') as token:
        pickle.dump(creds, token)

service = build('sheets', 'v4', credentials=creds)

# Call the Sheets API
sheet = service.spreadsheets()
result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range='A1:Z999').execute()     # Read the whole thing.
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

# Drop the room names from the spreadsheet
roomNames=[r.strip() for r in values[0]]
values=values[1:]

# Start building the participants and items databases (dictionaries)
participants={} # A dictionary keyed by a person's name containing a list of (time, room, item) tuples, each an item that that person is on.
items={}        # A dictionary keyed by item name containing a (time, room, people-list) tuple, where people-list is the list of people on the item
times=[]        # This is a list of times in spreadsheet order which should be in sorted order.

# When we find a row with data in column 0, we have found a new time.
rowIndex=0
while rowIndex < len(values):
    row=values[rowIndex]
    if len(row) == 0:   # Skip empty rows
        rowIndex+=1
        continue
    time=row[0].strip() # When a row has the first column filled, that element is the time of the item
    times.append(time)
    # Lookin at the rest of the row, there may be text in one or more of the room columns
    for roomIndex in roomIndexes:
        if roomIndex < len(row):    # Trailing empty cells have been truncated, so better check.
            if len(row[roomIndex]) > 0:     # So does the cell itself contain text?
                # This has to be an item name since it's a cell containing text in a row that starts with a timeand in a column that starts with a room
                itemName=row[roomIndex].strip()
                # If there are people scheduled for it, they will be in the next cell down
                peopleRow=rowIndex+1
                peopleList=[]
                if len(values)> peopleRow:  # Does peopleRow exist?
                    if len(values[peopleRow]) > roomIndex:  # Does it have enough columns
                        if len(values[peopleRow][roomIndex]) > 0: # Does it have anything in the right column?
                            people=values[peopleRow][roomIndex].split(",")  # Get a list of people
                            for person in people:
                                person=person.strip()
                                if len(person) > 0:     # If there's anything left, add this item to that person's entry
                                    if person not in participants.keys():   # If this is the first time we've encountered this person, create an empty entry.
                                        participants[person]=[]
                                    participants[person].append((time, roomNames[roomIndex], itemName))     # And append a tuple with the time, room, and item name
                                    peopleList.append(person)
                items[itemName]=(time, roomNames[roomIndex], peopleList)
    rowIndex+=2 # Skip both rows

# Print the items by people with time list
# Get a list of the program participants (the keys of the  participants dictionary) sorted by the last token in the name (which will usually be the last name)
partlist=sorted(participants.keys(), key=lambda x: x.split(" ")[-1])
txt=open("People with items by time.txt", "w")
for person in partlist:
    print("", file=txt)
    print(person, file=txt)
    for item in participants[person]:
        print("    "+item[0]+": "+item[2], file=txt)
txt.close()

# Now the raw text for the pocket program
txt=open("Pocket program.txt", "w")
for time in times:
    print("\n"+time, file=txt)
    for room in roomNames:
        # Now search for the program item and people list for this slot
        for itemName in items.keys():
            item=items[itemName]
            if item[0] == time and item[1] == room:
                print("   "+room+":  "+itemName, file=txt)
                if item[2] is not None and len(item[2]) > 0:
                    print("            "+", ".join(item[2]), file=txt)
txt.close()