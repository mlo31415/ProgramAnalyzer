import pickle
import os.path
import difflib
import docx
import math
import re as RegEx
from docx.shared import Pt
from docx.shared import Inches
from docx import text
from docx.text import paragraph
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from ScheduleItem import ScheduleItem
from Item import Item

#*************************************************************************************************
#*************************************************************************************************
# Miscellaneous helper functions

# Generate the name of a person stripped if any "(M)" or "(m)" flags
def RemoveModFlag(s: str):
    return s.replace("(M)", "").replace("(m)", "").strip()

# Is this person's name flagged as a moderator?
def IsModerator(s: str):
    return s != RemoveModFlag(s)

# Delete a file, ignoring any errors
# We do this because of as-yet not understood failures to delete files
def SafeDelete(fn: str):
    try:
        os.remove(fn)
    except:
        return

# Convert a text date string to numeric
def TextToNumericTime(s: str):
    global gDayList
    # The date string is of the form Day Hour AM/PM or Day Noon
    day=""
    hour=""
    minutes=""
    suffix=""

    m=RegEx.match(r"^([A-Za-z]+)\s*([0-9]+)\s*([A-Za-z]+)$", s)     # <day> <hr> <am/pm/noon/etc>
    if m is not None:
        day=m.groups()[0]
        hour=m.groups()[1]
        suffix=m.groups()[2]
    else:
        m=RegEx.match(r"^([A-Za-z]+)\s*([0-9]+):([0-9]+)\s*([A-Za-z]+)$", s)    # <day> <hr>:<min> <am/pm/noon/etc>
        if m is not None:
            day=m.groups()[0]
            hour=m.groups()[1]
            minutes=m.groups()[2]
            suffix=m.groups()[3]
        else:
            m=RegEx.match(r"^([A-Za-z]+)\s*([A-Za-z]+)$", s)    # <day> <am/pm/noon/etc>
            if m is not None:
                day=m.groups()[0]
                suffix=m.groups()[1]
            else:
                print("Can't interpret time: '"+s+"'")

    d=gDayList.index(day)
    h=0
    if hour != "":
        h=int(hour)
    if minutes != "":
        h=h+int(minutes)/60
    if suffix.lower() == "pm":
        h=h+12
    elif suffix.lower() == "noon":
        h=12
    elif suffix.lower() == "midnight":
        h=24

    #print("'"+s+"'  --> day="+day+"  hour="+hour+"  minutes="+minutes+"  suffix="+suffix+"   --> d="+str(d)+"  h="+str(h)+"  24*d+h="+(str(24*d+h))+"  --> "+NumericToTextDayTime(24*d+h))
    return 24*d+h

# Convert a numeric daytime to text
# The input time is a floating point number of hours since the start of the 1st day of the convention
def NumericToTextDayTime(f: float):
    global gDayList
    d=math.floor(f/24)  # Compute the day number
    return gDayList[int(d)] + " " + NumericToTextTime(f)


def NumericToTextTime(f: float):
    d=math.floor(f/24)  # Compute the day number
    f=f-24*d
    isPM=f>12           # AM or PM?
    if isPM:
        f=f-12
    h=math.floor(f)     # Get the hour
    f=f-h               # What's left is the fractional hour

    if h == 12:         # Handle noon and midnight specially
        if isPM:
            return "Midnight"
        else:
            return "Noon"

    if h == 0 and f != 0:
        numerictime="12:"+str(math.floor(60*f))     # Handle the special case of times after noon but before 1
    else:
        numerictime=str(h) + ("" if f == 0 else ":" + str(math.floor(60*f)))

    return numerictime + ("pm" if isPM else "am")


# Return the name of the day corresponding to a numeric time
def NumericTimeToDayString(f: float):
    global gDayList
    d=math.floor(f/24)  # Compute the day number
    return gDayList[int(d)]


def FmtLen(val: list):
    if val is None:
        return "0"
    return str(len(val))



#*************************************************************************************************
#*************************************************************************************************
# MAIN
# Read and analyze the spreadsheet

# Create the reports subfolder if none exists
if not os.path.exists("reports"):
    os.mkdir("reports")

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
scheduleCells = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range='Schedule!A1:Z1999').execute().get('values', [])     # Read the whole thing.
if not scheduleCells:
    print("Can't locate scheduleCells tab in spreadsheet")
    raise(ValueError, "No scheduleCells found")

precisCells = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range='Precis!A1:Z999').execute().get('values', [])     # Read the whole thing.
if not precisCells:
    print("Can't locate precisCells tab in spreadsheet")
    raise(ValueError, "No precisCells found")
precisCells=[p for p in precisCells if len(p) > 0 and "".join(p)[0] != "#"]      # Drop blank lines and lines with a "#" alone in column 1.if not precisCells:

peopleCells = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range='People!A1:Z999').execute().get('values', [])     # Read the whole thing.
if not peopleCells:
    print("Can't locate peopleCells tab in spreadsheet")
    raise(ValueError, "No peopleCells found")
peopleCells=[p for p in peopleCells if len(p) > 0 and "".join(p)[0] != "#"]      # Drop blank lines and lines with a "#" alone in column 1.

parameterCells = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range='Controls!A1:Z999').execute().get('values', [])     # Read the whole thing.
if not parameterCells:
    print("Can't locate parameterCells tab in spreadsheet")
    raise(ValueError, "No parameterCells found")
parameterCells=[p for p in parameterCells if len(p) > 0 and "".join(p)[0] != "#"]      # Drop blank lines and lines with a "#" alone in column 1.

# Read parameters from the Control sheet
startingDay="Friday"
for row in parameterCells:
    if row[0] == "Starting day":
        if len(row) > 1:
            startingDay=row[1].strip()
            startingDay=startingDay[0].upper()+startingDay.lower()[1:]  # Force the capitalization to be right
# Reorganize the dayList so it starts with our starting day. It's extra-long so that clipping days from the front will still leave a full week.
gDayList=["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]
if startingDay not in gDayList:
    print("Can't interpret starting day='"+startingDay+"'.  Will use 'Friday'")
    startingDay="Friday"
i=gDayList.index(startingDay)
gDayList=gDayList[i:]

# We're done with reading the spreadsheet. Now analyze the data.
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

# Get the room names which are in the first row of the scheduleCells tab
gRoomNames=[r.strip() for r in scheduleCells[0]]

if len(gRoomNames) == 0 or len(roomIndexes) == 0:
    print("Room names line is blank.")

# Start reading ths spreadsheet and building the participants and items databases (dictionaries)
gSchedules={}   # A dictionary keyed by a person's name containing a list of (time, room, item, moderator) tuples, each an item that that person is on.
gItems={}       # A dictionary keyed by item name containing a (time, room, people-list, moderator) tuple, where people-list is the list of people on the item
gTimes=[]       # A list of times found in the spreadsheet.

#.......
# Add an item with a list of people, and add the item to each of the persons
def AddItemWithPeople(time: float, roomName: str, itemName: str, plistText: str):
    global gSchedules
    global gItems

    plist=plistText.split(",")  # Get the people as a list
    plist=[p.strip() for p in plist]  # Remove excess spaces
    plist=[p for p in plist if len(p) > 0]
    modName=None
    peopleList=[]
    for person in plist:  # For each person listed on this item
        if IsModerator(person):
            modName=person=RemoveModFlag(person)
        if person not in gSchedules.keys():  # If this is the first time we've encountered this person, create an empty entry.
            gSchedules[person]=[]
        gSchedules[person].append(ScheduleItem(PersonName=person, Time=time, Room=roomName, ItemName=itemName, IsMod=(person == modName)))  # And append a tuple with the time, room, item name, and moderator flag
        peopleList.append(person)
    # And add the item with its list of people to the items table.
    if itemName in gItems:  # If the item's name is already in use, add a uniquifier of room+day/time
        itemName=itemName+"  {"+roomName+" "+NumericToTextDayTime(time)+"}"
    gItems[itemName]=Item(Name=itemName, Time=time, Room=roomName, People=peopleList, ModName=modName)


#.......
# Add an item with a list of people, and add the item to each of the persons
def AddItemWithoutPeople(time: float, roomName: str, itemName: str):
    global gItems
    if itemName in gItems:  # If the item's name is already in use, add a uniquifier of room+day/time
        itemName=itemName+"  {"+roomName+" "+NumericToTextDayTime(time)+"}"
    gItems[itemName]=Item(Name=itemName, Time=time, Room=roomName, People=None, ModName=None)

#.......
# Code to process a set of time and people rows.
def ProcessRows(timeRow, peopleRow):
    # Get the time from the timerow and add it to gTimes
    time=TextToNumericTime(timeRow[0])
    if time not in gTimes:
        gTimes.append(time)     # We want to allow duplicate time rows, just-in-case

    # Looking at the rest of the row, there may be text in one or more of the room columns
    for roomIndex in roomIndexes:
        roomName=gRoomNames[roomIndex]
        if roomIndex < len(timeRow):    # Trailing empty cells have been truncated, so better check.
            if len(timeRow[roomIndex]) > 0:     # So does the cell itself contain text?
                # This has to be an item name since it's a cell containing text in a row that starts with a time and in a column that starts with a room
                itemName=timeRow[roomIndex]

                # Does a row indexed by peopleRowIndex exist in the spreadsheet? Does it have enough columns? Does it have anything in the correct column?
                if peopleRow is not None and len(peopleRow) > roomIndex and len(peopleRow[roomIndex]) > 0:
                    # We indicate items which go for an hour, but have some people in one part and some in another using a special notation in the people list.
                    # Robert A. Heinlein, [0.5] John W. Campbell puts RAH on the hour and JWC a half-hour later.
                    # There is much messiness in this.
                    # We look for the [##] in the people list.  If we find it, we divide the people list in half and create two items with separate plists.
                    plistText=peopleRow[roomIndex]
                    r=RegEx.match("(.*)\[([0-9.]*)\](.*)", plistText)
                    if r is None:
                        AddItemWithPeople(time, roomName, itemName, plistText)
                    else:
                        plist1=r.groups()[0].strip()
                        deltaT=r.groups()[1].strip()
                        plist2=r.groups()[2].strip()
                        AddItemWithPeople(time, roomName, itemName, plist1)
                        newTime=time+float(deltaT)
                        if newTime not in gTimes:
                            gTimes.append(newTime)
                        # This second instance will need to have a distinct item name, so add {#2} to the item name
                        AddItemWithPeople(newTime, roomName, itemName+" {#2}", plist2)
                else:   # We have an item with no people on it.
                    AddItemWithoutPeople(time, roomName, itemName)


# Now process the schedule row by row
# When we find a row with data in column 0, we have found a new time.
# A time row contains items.
# A time row will normally be followed by a people row containing the participants for those items
rowIndex=1  # We skip the first row which contains room names
timeRow=None
while rowIndex < len(scheduleCells):
    row=[c.strip() for c in scheduleCells[rowIndex]]  # Get the next row as a list of cells. Strip off leading and trailing blanks for each cell.
    if len(row) == 0:   # Ignore empty rows
        rowIndex+=1
        continue

    # Skip rows where the first character is a "#"
    if "".join(row)[0] == "#":
        rowIndex+=1
        continue

    # We have a non-empty row. It can be a time row or a people row.
    if len(row[0]) > 0:     # Is it a time row?
        if timeRow is None:     # Is it a time row without a pending time row?
            # Save it and go to the next row
            timeRow=row
            rowIndex+=1
            continue

        if timeRow is not None: # If it's a time row, and we already have a time row, then that first time row has no people.  Process the first row, save the second and move on.
            # Process time row without people row
            ProcessRows(timeRow, None)
            timeRow=row
            rowIndex+=1
            continue
    else:   # Then it must be a people row
        if timeRow is not None:      # Is it a people row following a time row?
            ProcessRows(timeRow, row)
            rowIndex+=1
            timeRow=None
            continue

        if timeRow is None:    # Is it a people row that doesn't follow a time row?
            print("Error reading schedule: Row "+str(rowIndex+1)+" is a people row; we were expecting a time row.")     # +1 because the spreadsheet's row-numbering is 1-based
            print("   row="+" ".join(row))
            i=0
            rowIndex+=1

if timeRow is not None:
    ProcessRows(timeRow, None)

# Make sure times are sorted into ascending order.
# The simple sort works because the times are stored as numeric hours since start of first day.
gTimes.sort()

#******
# Analyze the Precis cells and add the information to the
# The first row is column labels. So ignore it.
precisCells=precisCells[1:]

# The rest of the rows of the tab is pairs title:precis.
count=0
fname=os.path.join("reports", "Diag - precis without items.txt")
txt=open(fname, "w")
print("Precis without corresponding items:", file=txt)
for row in precisCells:
    row=[r.strip() for r in row]    # Get rid of leading and trailing blanks
    if len(row[0]) > 0 and len(row[1]) > 0: # If both the item name and the precis exist, store them in the precis table.
        itemname=row[0]
        if itemname not in gItems.keys():
            count+=1
            print("   "+itemname, file=txt)
        else:
            gItems[itemname].Precis=row[1]
if count == 0:
    print("    None found", file=txt)
txt.close()

#******
# Analyze the People cells

# Step 1 is to find the column labels.
# They are in the first non-empty row.
firstNonEmptyRow=0
while firstNonEmptyRow < len(peopleCells):
    if len(peopleCells[firstNonEmptyRow]) > 0:      # Rely on Googledocs truncating of trailing empty cells so that a blank line has no cells in it.
        break
    firstNonEmptyRow+=1

# The first non-empty row is column labels.  Read them and identify the Fname, Lname, Email, and Response columns
fnameCol=None
lnameCol=None
emailCol=None
responseCol=None
for i in range(0, len(peopleCells[firstNonEmptyRow])):
    cell=peopleCells[firstNonEmptyRow][i].lower()
    if cell == "fname":
        fnameCol=i
    if cell == "lname":
        lnameCol=i
    if cell == "email":
        emailCol=i
    if cell == "response":
        responseCol=i

#TODO: Need some sort of error report if the fname, lname, or response is missing
# We'll combine the first and last names to create a full name like is used elsewhere.
peopleTable={}
for i in range(firstNonEmptyRow+1, len(peopleCells)):
    if len(peopleCells) == 0:   # Skip empty rows
        continue
    row=[r.strip() for r in peopleCells[i]]    # Get rid of leading and trailing blanks in each cell
    fname=""
    if fnameCol < len(row):
        fname=row[fnameCol]
    lname=""
    if lnameCol < len(row):
        lname=row[lnameCol]
    fullname=None
    if len(fname) > 0 and len(lname) > 0:   # Gotta handle Ctein!
        fullname=fname+" "+lname
    elif len(fname) > 0:
        fullname=fname
    elif len(lname) > 0:
        fullname=lname

    email=""
    if emailCol < len(row):
        email=row[emailCol]
    response=""
    if responseCol < len(row):
        response=row[responseCol]

    if fullname is not None:    # TODO, and what if it is?
        peopleTable[fullname]=email, response.lower()       # Store the email and response as a tuple in the entry indexed by the full name


#*************************************************************************************************
#*************************************************************************************************
# Generate reports
# The first reports are all error reports or checking reports

#******
# Check for people in the schedule who are not in the people tab
fname=os.path.join("reports", "Diag - People in schedule without email.txt")
txt=open(fname, "w")
print("People who are scheduled but lack email address:", file=txt)
print("(Note that these may be due to spelling differences, use of initials, etc.)", file=txt)
count=0
for personname in gSchedules.keys():
    if personname not in peopleTable.keys():
        count+=1
        print("   "+personname, file=txt)
if count == 0:
    print("    None found", file=txt)
txt.close()


#******
# Check for people who are scheduled opposite themselves
fname=os.path.join("reports", "Diag - People scheduled against themselves.txt")
txt=open(fname, "w")
print("People who are scheduled to be in two places at the same time", file=txt)
count=0
for personname in gSchedules.keys():
    pSched=gSchedules[personname] # pSched is a person's schedule, which is a list of (time, room, item) tuples
    # Sort pSched by time, then look for duplicate times
    pSched.sort(key=lambda x: x.Time)
    last=ScheduleItem()
    for part in pSched:
        if part.Time == last:
            print(personname+": "+NumericToTextDayTime(last[0])+": "+last[1]+" and also "+part.Room, file=txt)
            count+=1
        last=part
if count == 0:
    print("    None found", file=txt)
txt.close()


#******
# Now look for similar name pairs
# First we make up a list of all names that appear in any tab
names=set()
names.update(gSchedules.keys())
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

#*******
# Print the People with items by time report
# Get a list of the program participants (the keys of the  participants dictionary) sorted by the last token in the name (which will usually be the last name)
sortedallpartlist=sorted(gSchedules.keys(), key=lambda x: x.split(" ")[-1])
fname=os.path.join("reports", "People with items by time.txt")
SafeDelete(fname)
txt=open(fname, "w")
for personname in sortedallpartlist:
    print("\n"+personname, file=txt)
    for schedItem in gSchedules[personname]:
        print("    "+NumericToTextDayTime(schedItem.Time)+": "+schedItem.DisplayName+" ["+schedItem.Room+"]"+(" (moderator)" if schedItem.IsMod else ""), file=txt)
txt.close()

#*******
# Print the Items with people by time report
# Get a list of the program participants (the keys of the  participants dictionary) sorted by the last token in the name (which will usually be the last name)
fname=os.path.join("reports", "Items with people by time.txt")
SafeDelete(fname)
txt=open(fname, "w")
for time in gTimes:
    for room in gRoomNames:
        # Now search for the program item and people list for this slot
        for itemName, item in gItems.items():
            if item.Time == time and item.Room == room:
                print(NumericToTextDayTime(time)+", "+room+": "+itemName+"   "+item.DisplayPlist(), file=txt)
                if item.Precis is not None:
                    print("     "+item.Precis, file=txt)
txt.close()


#*******
# Print the program participant's schedule report
# Get a list of the program participants (the keys of the  participants dictionary) sorted by the last token in the name (which will usually be the last name)
fname=os.path.join("reports", "Program participant schedules.txt")
SafeDelete(fname)
txt=open(fname, "w")
for personname in sortedallpartlist:
    print("\n\n********************************************", file=txt)
    print(personname, file=txt)
    for schedItem in gSchedules[personname]:
        print("\n"+NumericToTextDayTime(schedItem.Time)+": "+schedItem.DisplayName+" ["+schedItem.Room+"]"+(" (moderator)" if schedItem.IsMod else ""), file=txt)
        item=gItems[schedItem.ItemName]
        print("Participants: "+item.DisplayPlist(), file=txt)
        if item.Precis is not None:
            print("Precis: "+item.Precis, file=txt)
txt.close()


#******
# Report on the number of people/item
fname=os.path.join("reports", "Items' people counts.txt")
SafeDelete(fname)
txt=open(fname, "w")
print("List of number of people scheduled on each item\n\n", file=txt)
for itemname, item in gItems.items():
    print(NumericToTextDayTime(item.Time)+" "+item.Name+": "+FmtLen(item.People), file=txt)
txt.close()

#******
# Flag items with a suspiciously small number of people on them
fname=os.path.join("reports", "Diag - Items with unexpectedly low number of participants.txt")
SafeDelete(fname)
txt=open(fname, "w")
print("List of non-readings and KKs with fewer than 3 people on them\n\n", file=txt)
found=False
for itemname, item in gItems.items():
    if item.People is not None and len(item.People) >= 3:
        continue
    if item.Name.find("Reading") > -1 or item.Name.find("KK") > -1 or item.Name.find("Kaffe") > -1 or item.Name.find("Autograph") > -1:
        continue
    print(NumericToTextDayTime(item.Time)+" "+item.Name+": "+FmtLen(item.People), file=txt)
    found=True
if not found:
    print("None found")
txt.close()


#******
# Flag items missing a moderator or a precis
fname=os.path.join("reports", "Diag - Items missing a moderator.txt")
SafeDelete(fname)
txt=open(fname, "w")
print("List of non-readings and KKs with no moderator\n\n", file=txt)
found=False
for itemname, item in gItems.items():
    if item.Name.find("Reading") > -1 or item.Name.find("KK") > -1 or item.Name.find("Kaffe") > -1 or item.Name.find("Autograph") > -1:
        continue
    if item.ModName is not None:
        continue
    print(NumericToTextDayTime(item.Time)+" "+item.Name+": "+FmtLen(item.People), file=txt)
    found=True
if not found:
    print("None found")
txt.close()

fname=os.path.join("reports", "Diag - Items missing a precis.txt")
SafeDelete(fname)
txt=open(fname, "w")
print("List of non-readings and KKs with no precis\n\n", file=txt)
found=False
for itemname, item in gItems.items():
    if item.Name.find("Reading") > -1 or item.Name.find("KK") > -1 or item.Name.find("Kaffe") > -1 or item.Name.find("Autograph") > -1:
        continue
    if item.Precis is not None and len(item.Precis) > 0:
        continue
    print(NumericToTextDayTime(item.Time)+" "+item.Name+": "+FmtLen(item.People), file=txt)
    found=True
if not found:
    print("None found")
txt.close()

#******
# Report on the number of items/person
# Include all people in the people tab, even those with no items
fname=os.path.join("reports", "Peoples' item counts.txt")
SafeDelete(fname)
txt=open(fname, "w")
print("List of number of items each person is scheduled on\n\n", file=txt)
for personname in peopleTable:
    if personname in gSchedules.keys():
        print(personname+": "+str(len(gSchedules[personname]))+("" if peopleTable[personname][1] == "y" else " not confirmed"), file=txt)
    else:
        if peopleTable[personname][1] == "y":
            print(personname+": coming, but not scheduled", file=txt)
txt.close()


#******
# Create a docx and a .txt version for the pocket program
# Note that we're generating two files at once here.
def AppendParaToDoc(doc: docx.Document, txt: str, bold=False, italic=False, size=14, indent=0.0, font="Calibri"):
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

def AppendTextToPara(para: docx.text.paragraph.Paragraph, txt: str, bold: bool=False, italic: bool=False, size: float=14, indent: float=0.0, font: str="Calibri"):
    run=para.add_run(txt)
    run.bold=bold
    run.italic=italic
    runfont=run.font
    runfont.name=font
    runfont.size=Pt(size)
    para.paragraph_format.left_indent=Inches(indent)
    para.paragraph_format.line_spacing=1
    para.paragraph_format.space_after=0

doc=docx.Document()
fname=os.path.join("reports", "Pocket program.txt")
SafeDelete(fname)
txt=open(fname, "w")
AppendParaToDoc(doc, "Schedule", bold=True, size=24)
print("Schedule", file=txt)
for time in gTimes:
    AppendParaToDoc(doc, "")
    AppendParaToDoc(doc, NumericToTextDayTime(time), bold=True)
    print("\n"+NumericToTextDayTime(time), file=txt)
    for room in gRoomNames:
        # Now search for the program item and people list for this slot
        for itemName, item in gItems.items():
            if item.Time == time and item.Room == room:
                para=doc.add_paragraph()
                AppendTextToPara(para, room+": ", italic=True, size=12, indent=0.3)
                AppendTextToPara(para, item.DisplayName, size=12, indent=0.3)
                print("   "+room+":  "+item.DisplayName, file=txt)   # Print the room and item name
                if item.People is not None and len(item.People) > 0:            # And the item's people list
                    plist=item.DisplayPlist()
                    AppendParaToDoc(doc, plist, size=12, indent=0.6)
                    print("            "+plist, file=txt)
                if item.Precis is not None:
                    AppendParaToDoc(doc, item.Precis, italic=True, size=12, indent=0.6)
                    print("            "+item.Precis, file=txt)
fname=os.path.join("reports", "Pocket program.docx")
doc.save(fname)
txt.close()


#******
# Generate web pages, one for each day.
day=""
f=None
for time in gTimes:
    # We generate a separate report for each day
    d=NumericTimeToDayString(time)
    if d != day:
        # Close the old file, if any
        if f is not None:
            f.write('</font></table>\n')
            # Read and append the footer
            with open("control-WebpageFooter.txt", "r") as f2:
                f.writelines(f2.readlines())
            f.close()
            f=None
        # Open the new one
        day=d
        fname=os.path.join("reports", "Schedule - "+day+".html")
        SafeDelete(fname)
        f=open(fname, "w")
        with open("control-WebpageHeader.txt", "r") as f2:
            f.writelines(f2.readlines())
        f.write("<h2>"+day+"</h2>\n")
        f.write('<table border="0" cellspacing="0" cellpadding="2">\n')

    f.write('<tr><td colspan="3">')
    f.write('<p class="time">' + NumericToTextTime(time) + '</p>')
    f.write('</td></tr>\n')
    for room in gRoomNames:
        # Now search for the program item and people list for this slot
        for itemName, item in gItems.items():
            if item.Time == time and item.Room == room:
                f.write('<tr><td width="40">&nbsp;</td><td colspan="2">')   # Two columns, the first 40 pixes wide and empty
                f.write('<p><span class="room">' + room +': </span><span class="item">' + item.DisplayName +'</span></p>')
                f.write('</td></tr>')
                if item.People is not None and len(item.People) > 0:            # And the item's people list
                    f.write('<tr><td width="40">&nbsp;</td><td width="40">&nbsp;</td><td width="600">')     # Three columns, the first two 40 pixes wide and empty; the third 600 pixels wide
                    f.write('<p><span class="people">'+item.DisplayPlist()+'</span></p>')
                    f.write('</td></tr>\n')
                if item.Precis is not None:
                    f.write('<tr><td width="40">&nbsp;</td><td width="40">&nbsp;</td><td width="600">')     # Same
                    f.write('<p><span class="precis">'+item.Precis+'</span></p>')
                    f.write('</td></tr>\n')
if f is not None:
    # Read and append the footer
    f.write('</table>\n')
    with open("control-WebpageFooter.txt", "r") as f2:
        f.writelines(f2.readlines())
    f.close()


#******
# Do the room signs.  They'll go in reports/rooms/<name>.docx
# Create the roomsigns subfolder if none exists
path=os.path.join("reports", "roomsigns")
if not os.path.exists(path):
    os.mkdir(path)
for room in gRoomNames:
    inuse=False  # Make sure that this room is actually in use
    if len(room.strip()) == 0:
        continue
    doc=docx.Document()
    AppendParaToDoc(doc, room, bold=True, size=32)  # Room name at top
    for time in gTimes:
        for itemName in gItems.keys():
            item=gItems[itemName]
            if item.Time == time and item.Room == room:
                inuse=True
                AppendParaToDoc(doc, "")    # Skip a line
                para=doc.add_paragraph()
                AppendTextToPara(para, NumericToTextDayTime(item.Time)+":  ", bold=True)   # Add the time in bold followed by the item's title
                AppendTextToPara(para, item.DisplayName)
                AppendParaToDoc(doc, item.DisplayPlist(), italic=True, indent=0.5)        # Then, on a new line, the people list in italic
    fname=os.path.join(path, room.replace("/", "-")+".docx")
    SafeDelete(fname)
    if inuse:
        doc.save(fname)
