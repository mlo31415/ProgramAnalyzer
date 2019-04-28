import pickle
import os.path
import difflib
import docx
from docx.shared import Pt
from docx.shared import Inches
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request


# Generate the display-name of an item. (Remove any text following the first "{")
def ItemDisplayName(s: str):
    loc=s.find("{")
    if loc > 0:
        return s[:loc-1]
    return s

def ItemDisplayPlist(item: tuple):
    s=""
    for person in item[2]:
        s=s + (", " if len(s) > 0 else "") + person + (" (M)" if person == item[3] else "")
    return s

# Generate the name of a person stripped if any "(M)" or "(m)" flags
def RemoveModFlag(s: str):
    return s.replace("(M)", "").replace("(m)", "").strip()

# Is this person's name flagged as a moderator?
def IsModerator(s: str):
    return s != RemoveModFlag(s)

# Delete a file, ignoring any errors
# We do this because of as-yet not understood failures to delete files
def SafeDelete(fn):
    try:
        os.remove(fn)
    except:
        return

#*************************************************************************************************
#*************************************************************************************************
# MAIN
# Read and analyze the spreadsheet

credentials = None
# The file token.pickle stores the user's access and refresh tokens, and is
# created automatically when the authorization flow completes for the first time.
# Pickle is a scheme for serializing data to disk and retrieving it
if os.path.exists('token.pickle'):
    with open('token.pickle', 'rb') as token:
        credentials = pickle.load(token)

# If there are no (valid) credentials available, let the user log in.
if not credentials or not credentials.valid:
    if credentials and credentials.expired and credentials.refresh_token:
        credentials.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file('credentials.json', 'https://www.googleapis.com/auth/spreadsheets.readonly')
        credentials = flow.run_local_server()
    # Save the credentials for the next run
    with open('token.pickle', 'wb') as token:
        pickle.dump(credentials, token)

service = build('sheets', 'v4', credentials=credentials)

# Call the Sheets API to load the various tabs of the spreadsheet
sheet = service.spreadsheets()
SPREADSHEET_ID ='1UjHSw-R8dLNFGctUhIQiPr58aAAfBedGznJEN2xBn7o'  # This is the ID of the specific spreadsheet we're reading
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
# The first row of the spreadsheet is the list of rooms.
# Make a list of room names and room column indexes
roomIndexes=[]
for i in range(0, len(scheduleCells[0])):
    if scheduleCells[0][i] is None:
        break
    if len(scheduleCells[0][i]) > 0:
        roomIndexes.append(i)

# Drop the room names from the spreadsheet leaving just the schedules
roomNames=[r.strip() for r in scheduleCells[0]]
scheduleCells=scheduleCells[1:]

# Start reading ths spreadsheet and building the participants and items databases (dictionaries)
participants={} # A dictionary keyed by a person's name containing a list of (time, room, item) tuples, each an item that that person is on.
items={}        # A dictionary keyed by item name containing a (time, room, people-list, moderator) tuple, where people-list is the list of people on the item
times=[]        # This is a list of times in spreadsheet order which should be in sorted order.

# When we find a row with data in column 0, we have found a new time.
rowIndex=0
while rowIndex < len(scheduleCells):
    row=[c.strip() for c in scheduleCells[rowIndex]]  # Get just the one row as a list of cells. Strip off leading and trailing blanks for each cell.
    if len(row) == 0:   # Ignore empty rows
        rowIndex+=1
        continue

    time=row[0] # When a row has the first column filled, that element is the time of the item.  If the spreadsheet is well-formed, the next non-blank line is a time line
    times.append(time)
    # Looking at the rest of the row, there may be text in one or more of the room columns
    for roomIndex in roomIndexes:
        if roomIndex < len(row):    # Trailing empty cells have been truncated, so better check.
            if len(row[roomIndex]) > 0:     # So does the cell itself contain text?
                # This has to be an item name since it's a cell containing text in a row that starts with a time and in a column that starts with a room
                itemName=row[roomIndex]
                modName=""
                # If there are people scheduled for it, they will be in the next cell down
                peopleRowIndex=rowIndex+1
                peopleList=[]
                if len(scheduleCells)> peopleRowIndex:  # Does a row indexed by peopleRowIndex exist in the spreadsheet?
                    if len(scheduleCells[peopleRowIndex]) > roomIndex:  # Does it have enough columns?
                        if len(scheduleCells[peopleRowIndex][roomIndex]) > 0: # Does it have anything in the right column?
                            people=scheduleCells[peopleRowIndex][roomIndex].split(",")  # Get a list of people
                            people=[p.strip() for p in people]
                            for person in people:       # For each person listed on this item
                                if len(person) > 0:     # Is it's not empty...
                                    if IsModerator(person):
                                        modName=person=RemoveModFlag(person)
                                    if person not in participants.keys():   # If this is the first time we've encountered this person, create an empty entry.
                                        participants[person]=[]
                                    participants[person].append((time, roomNames[roomIndex], itemName, person == modName))     # And append a tuple with the time, room, item name, and moderator flag
                                    peopleList.append(person)
                items[itemName]=(time, roomNames[roomIndex], peopleList, modName)
    rowIndex+=2 # Skip both rows

#******
# Analyze the Precis cells
# The first row is column labels. So ignore it.
precisCells=precisCells[1:]

# The rest of the tab is pairs title:precis.
precis={}
for row in precisCells:
    row=[r.strip() for r in row]    # Get rid of leading and trailing blanks
    if len(row[0]) > 0 and len(row[1]) > 0: # If both the item name and the precis exist, store them in the precis table.
        precis[row[0]]=row[1]

#******
# Analyze the People cells
# The first row is column labels. So ignore it.
peopleCells=peopleCells[1:]

# The first two columns are first name and last name.  The third column is email
# We'll combine the first and last names to create a full name like is used elsewhere.
peopleTable={}
for row in peopleCells:
    row=[r.strip() for r in row]    # Get rid of leading and trailing blanks
    if len(row) > 2 and len(row[0]) > 0 and len(row[1]) > 0 and len(row[2]) > 0:
        peopleTable[row[0]+" "+row[1]]=row[2]       # Store the email in the entry indexed by the full name


#*************************************************************************************************
#*************************************************************************************************
# Generate reports
# The first reports are all error reports or checking reports

# Create the reports subfolder if none exists
if not os.path.exists("reports"):
    os.mkdir("reports")

# Print a list of precis without corresponding items and items without precis
fname=os.path.join("reports", "Diag - Precis without items and items without precis.txt")
txt=open(fname, "w")
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


#******
# Check for people in the schedule who are not in the people tab
fname=os.path.join("reports", "Diag - People in schedule without email.txt")
txt=open(fname, "w")
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


#******
# Check for people who are scheduled opposite themselves
fname=os.path.join("reports", "Diag - People scheduled against themselves.txt")
txt=open(fname, "w")
print("People who are scheduled to be in two places at the same time", file=txt)
count=0
for person in participants.keys():
    pSched=participants[person] # pSched is a person's schedule, which is a list of (time, room, item) tuples
    # Sort pSched by time, then look for duplicate times
    pSched.sort(key=lambda x: x[0])
    last=(0,0,0)
    for it in pSched:
        if it[0] == last[0]:
            print(person+": "+last+"    "+it)
            count+=1
        last=it
if count == 0:
    print("    None found", file=txt)
txt.close()


#******
# Now look for similar name pairs
# First we make up a list of all names that appear in any tab
names=set()
names.update(participants.keys())
names.update(peopleTable.keys())
similarNames=[]
for p1 in names:
    for p2 in names:
        if p1 < p2:
            rat=difflib.SequenceMatcher(a=p1, b=p2).ratio()
            if rat > .75:
                similarNames.append((p1, p2, rat))
similarNames.sort(key=lambda x: x[2], reverse=True)


fname=os.path.join("reports", "Diag - Disturbingly similar names.txt")
SafeDelete(fname)
if len(similarNames) > 0:
    txt=open(fname, "w")
    print("Names that are disturbingly similar:", file=txt)
    count=0
    for s in similarNames:
        print("   "+s[0]+"  &  "+s[1], file=txt)
        count+=1
    if count == 0:
        print("    None found", file=txt)
    txt.close()


#****************************************************
# Now do the content/working reports

# Print the People with items by time report
# Get a list of the program participants (the keys of the  participants dictionary) sorted by the last token in the name (which will usually be the last name)
partlist=sorted(participants.keys(), key=lambda x: x.split(" ")[-1])
fname=os.path.join("reports", "People with items by time.txt")
txt=open(fname, "w")
for person in partlist:
    print("", file=txt)
    print(person, file=txt)
    for item in participants[person]:
        print("    " + item[0] + ": " + ItemDisplayName(item[2]) + (" (moderator)" if item[3] else ""), file=txt)
txt.close()


def AppendParaToDoc(doc, txt: str, bold=False, italic=False, size=14, indent=0.0, font="Calibri"):
    para=doc.add_paragraph()
    run=para.add_run(txt)
    run.bold=bold
    run.italic=italic
    runfont=run.font
    runfont.name=font
    runfont.size=Pt(size)
    para.paragraph_format.left_indent=Inches(indent)
    para.paragraph_format.line_spacing=1
    para.paragraph_format.space_after=0

def AppendTextToPara(para, txt: str, bold=False, italic=False, size=14, indent=0.0, font="Calibri"):
    run=para.add_run(txt)
    run.bold=bold
    run.italic=italic
    runfont=run.font
    runfont.name=font
    runfont.size=Pt(size)
    para.paragraph_format.left_indent=Inches(indent)
    para.paragraph_format.line_spacing=1
    para.paragraph_format.space_after=0

# Create a docx and a .txt version for the pocket program
doc=docx.Document()
fname=os.path.join("reports", "Pocket program.txt")
txt=open(fname, "w")
AppendParaToDoc(doc, "Schedule", bold=True, size=24)
print("Schedule", file=txt)
for time in times:
    AppendParaToDoc(doc, "")
    AppendParaToDoc(doc, time, bold=True)
    print("\n"+time, file=txt)
    for room in roomNames:
        # Now search for the program item and people list for this slot
        for itemName in items.keys():
            item=items[itemName]
            if item[0] == time and item[1] == room:
                para=doc.add_paragraph()
                AppendTextToPara(para, room+": ", italic=True, size=12, indent=0.3)
                AppendTextToPara(para, ItemDisplayName(itemName), size=12, indent=0.3)
                print("   "+room+":  "+ItemDisplayName(itemName), file=txt)   # Print the room and item name
                if item[2] is not None and len(item[2]) > 0:            # And the item's people list
                    plist=ItemDisplayPlist(item)
                    AppendParaToDoc(doc, plist, size=12, indent=0.6)
                    print("            "+plist, file=txt)
                if itemName in precis.keys():
                    AppendParaToDoc(doc, precis[itemName], italic=True, size=12, indent=0.6)
                    print("            "+precis[itemName], file=txt)
fname=os.path.join("reports", "Pocket program.docx")
doc.save(fname)
txt.close()

#******
# Do the room signs.  They'll go in reports/rooms/<name>.docx
# Create the roomsigns subfolder if none exists
path=os.path.join("reports", "roomsigns")
if not os.path.exists(path):
    os.mkdir(path)
for room in roomNames:
    inuse=False  # Make sure that this room is actually in use
    if len(room.strip()) == 0:
        continue
    doc=docx.Document()
    AppendParaToDoc(doc, room, bold=True, size=32)  # Room name at top
    for time in times:
        for itemName in items.keys():
            item=items[itemName]
            if item[0] == time and item[1] == room:
                inuse=True
                AppendParaToDoc(doc, "")    # Skip a line
                para=doc.add_paragraph()
                AppendTextToPara(para, time+":  ", bold=True)   # Add the time in bold followed by the item's title
                AppendTextToPara(para, ItemDisplayName(itemName))
                plist=ItemDisplayPlist(item)
                AppendParaToDoc(doc, plist, italic=True, indent=0.5)        # Then, on a new line, the people list in italic
    fname=os.path.join(path, room+".docx")
    SafeDelete(fname)
    if inuse:
        doc.save(fname)
