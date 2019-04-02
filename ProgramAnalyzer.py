import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request


def ItemClean(s: str):
    loc=s.find("{")
    if loc > 0:
        return s[:loc-1]
    return s


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
scheduleCells = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range='Schedule!A1:Z999').execute().get('values', [])     # Read the whole thing.
if not scheduleCells:
    raise(ValueError, "No scheduleCells found")
precisCells = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range='Precis!A1:Z999').execute().get('values', [])     # Read the whole thing.
if not precisCells:
    raise(ValueError, "No precisCells found")
peopleCells = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range='People!A1:Z999').execute().get('values', [])     # Read the whole thing.
if not peopleCells:
    raise(ValueError, "No peopleCells found")


#******
# Analyze the Schedule cells
# The first row is the rooms.
# Make a list of room names and room column indexes
roomIndexes=[]
for i in range(0, len(scheduleCells[0])):
    if scheduleCells[0][i] is None:
        break
    if len(scheduleCells[0][i]) > 0:
        roomIndexes.append(i)

# Drop the room names from the spreadsheet
roomNames=[r.strip() for r in scheduleCells[0]]
scheduleCells=scheduleCells[1:]

# Start building the participants and items databases (dictionaries)
participants={} # A dictionary keyed by a person's name containing a list of (time, room, item) tuples, each an item that that person is on.
items={}        # A dictionary keyed by item name containing a (time, room, people-list) tuple, where people-list is the list of people on the item
times=[]        # This is a list of times in spreadsheet order which should be in sorted order.

# When we find a row with data in column 0, we have found a new time.
rowIndex=0
while rowIndex < len(scheduleCells):
    row=scheduleCells[rowIndex]
    if len(row) == 0:   # Skip empty rows
        rowIndex+=1
        continue
    time=row[0].strip() # When a row has the first column filled, that element is the time of the item
    times.append(time)
    # Lookin at the rest of the row, there may be text in one or more of the room columns
    for roomIndex in roomIndexes:
        if roomIndex < len(row):    # Trailing empty cells have been truncated, so better check.
            if len(row[roomIndex]) > 0:     # So does the cell itself contain text?
                # This has to be an item name since it's a cell containing text in a row that starts with a time and in a column that starts with a room
                itemName=row[roomIndex].strip()
                # If there are people scheduled for it, they will be in the next cell down
                peopleRow=rowIndex+1
                peopleList=[]
                if len(scheduleCells)> peopleRow:  # Does peopleRow exist?
                    if len(scheduleCells[peopleRow]) > roomIndex:  # Does it have enough columns
                        if len(scheduleCells[peopleRow][roomIndex]) > 0: # Does it have anything in the right column?
                            people=scheduleCells[peopleRow][roomIndex].split(",")  # Get a list of people
                            for person in people:
                                person=person.strip()
                                if len(person) > 0:     # If there's anything left, add this item to that person's entry
                                    if person not in participants.keys():   # If this is the first time we've encountered this person, create an empty entry.
                                        participants[person]=[]
                                    participants[person].append((time, roomNames[roomIndex], itemName))     # And append a tuple with the time, room, and item name
                                    peopleList.append(person)
                items[itemName]=(time, roomNames[roomIndex], peopleList)
    rowIndex+=2 # Skip both rows

#******
# Analyze the Precis cells
# The first row is column labels. So ignore it.
precisCells=precisCells[1:]

# The rest of the tab is pairs title:precis.
precis={}
for row in precisCells:
    if len(row[0].strip()) > 0 and len(row[1].strip()) > 0:
        precis[row[0].strip()]=row[1].strip()

#******
# Analyze the People cells
# The first row is column labels. So ignore it.
peopleCells=peopleCells[1:]

# the first two columns are first name and last name.  The third column is email
peopleTable={}
for row in peopleCells:
    if len(row) > 2 and len(row[0].strip()) > 0 and len(row[1].strip()) > 0 and len(row[2].strip()) > 0:
        peopleTable[row[0].strip()+" "+row[1].strip()]=row[2].strip()

# Create the reports folder if none exists
if not os.path.exists("reports"):
    os.mkdir("reports")

#**************************
# Generate reports

# Print a list of precis without corresponding items and items without precis
txt=open("reports/Diag - Precis without items and items without precis.txt", "w")
print("Items without precis:", file=txt)
count=0
for itemName in items.keys():
    if itemName not in precis.keys():
        count+=1
        print("   "+itemName, file=txt)
if count == 0:
    print("    None found", file=txt)

count=0
print("\n\nPrecis without items:", file=txt)
for itemName in precis.keys():
    if itemName not in items.keys():
        count+=1
        print("   "+itemName, file=txt)
if count == 0:
    print("    None found", file=txt)
txt.close()


# Check for people in the schedule who are not in the people tab
txt=open("reports/Diag - People in schedule without email.txt", "w")
print("People who are scheduled but lack email address:", file=txt)
print("(Note that these may be due to spelling differences, use of initials, etc.)", file=txt)
count=0
for person in participants.keys():
    if person not in peopleTable.keys():
        count+=1
        print("   "+person, file=txt)
if count == 0:
    print("    None found", file=txt)
txt.close()


# Print the items by people with time list
# Get a list of the program participants (the keys of the  participants dictionary) sorted by the last token in the name (which will usually be the last name)
partlist=sorted(participants.keys(), key=lambda x: x.split(" ")[-1])
txt=open("reports/People with items by time.txt", "w")
for person in partlist:
    print("", file=txt)
    print(person, file=txt)
    for item in participants[person]:
        print("    "+item[0]+": "+ItemClean(item[2]), file=txt)
txt.close()

# Now the raw text for the pocket program
txt=open("reports/Pocket program.txt", "w")
for time in times:
    print("\n"+time, file=txt)
    for room in roomNames:
        # Now search for the program item and people list for this slot
        for itemName in items.keys():
            item=items[itemName]
            if item[0] == time and item[1] == room:
                print("   "+room+":  "+ItemClean(itemName), file=txt)   # Print the room and item name
                if item[2] is not None and len(item[2]) > 0:            # And the item's people list
                    print("            "+", ".join(item[2]), file=txt)
                if itemName in precis.keys():
                    print("            "+precis[itemName], file=txt)

txt.close()
