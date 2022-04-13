from __future__ import annotations

from typing import Optional
from collections import defaultdict
from dataclasses import dataclass

import json
import os.path
import difflib
import docx
import wx
import re as RegEx
from docx.shared import Pt
from docx.shared import Inches
from docx import text
from docx.text import paragraph

from googleapiclient.discovery import build
from google.oauth2 import service_account
from googleapiclient.errors import HttpError

from HelpersPackage import PyiResourcePath, ReadListAsDict, MessageBox

from ScheduleItem import ScheduleItem
from Item import Item
from Log import Log, LogClose, LogError
import NumericTime


def main():
    # *************************************************************************************************
    # *************************************************************************************************
    # MAIN
    # Read and analyze the spreadsheet

    Log("Started")

    # Read the parameters.
    # This includes the names of the specific tabs to be used.
    parms=ReadListAsDict('parameters.txt')
    if len(parms) == 0:
        MessageBox("Can't open parameters.txt")
        app=wx.App()
        frame=wx.Frame(None, -1, 'win.py')
        frame.SetSize(0, 0, 200, 50)  # SetDimensions(0, 0, 200, 50)
        with wx.FileDialog(frame, "Open", "", "", "*.txt", wx.FD_OPEN|wx.FD_FILE_MUST_EXIST) as openFileDialog:
            ret=openFileDialog.ShowModal()
            if ret == wx.ID_CANCEL:
                exit(999)
            lst=openFileDialog.GetPath()
        Log(f"{lst} selected")
        parms=ReadListAsDict(lst)
        if len(parms) == 0:
            LogClose()
            exit(999)

    if len(parms["credentials"]) == 0:
        MessageBox("parameters.txt does not designate a credentials file")
        exit(999)

    with open(parms["credentials"]) as source:
        info=json.load(source)
        Log("Json read")

    if info is None:
        MessageBox("Json file is empty")
        exit(999)
    Log("credentials.txt read")

    # Create the reports subfolder if none exists
    parms.setdefault("reportdir", "Reports")
    reportsdir=parms["reportsdir"]
    if not os.path.exists(reportsdir):
        os.mkdir(reportsdir)
        Log(f"Reports directory {reportsdir} created ")

    credentials=service_account.Credentials.from_service_account_info(info)
    Log("Credentials established", Flush=True)

    service=build('sheets', 'v4', credentials=credentials)
    Log("Service established", Flush=True)

    # Call the Sheets API to load the various tabs of the spreadsheet
    googleSheets=service.spreadsheets()
    if len(parms["SheetID"]) == 0:
        MessageBox("parameters.txt does not designate a SheetID")
        exit(999)
    SPREADSHEET_ID=parms["SheetID"]  # This is the ID of the specific spreadsheet we're reading
    scheduleCells=ReadSheetFromTab(googleSheets, SPREADSHEET_ID, parms, "ScheduleTab")
    precisCells=ReadSheetFromTab(googleSheets, SPREADSHEET_ID, parms, "PrecisTab")
    peopleCells=ReadSheetFromTab(googleSheets, SPREADSHEET_ID, parms, "PeopleTab")
    parameterCells=ReadSheetFromTab(googleSheets, SPREADSHEET_ID, parms, "ControlTab")

    # Read parameters from the Control sheet
    startingDay="Friday"
    for row in parameterCells:
        if row[0] == "Starting day":
            if len(row) > 1:
                startingDay=row[1].strip()
                startingDay=startingDay[0].upper()+startingDay.lower()[1:]  # Force the capitalization to be right
    # Reorganize the dayList so it starts with our starting day. It's extra-long so that clipping days from the front will still leave a full week.
    if startingDay not in NumericTime.gDayList:
        LogError("Can't interpret starting day='"+startingDay+"'.  Will use 'Friday'")
        startingDay="Friday"
    i=NumericTime.gDayList.index(startingDay)
    NumericTime.gDayList=NumericTime.gDayList[i:]

    # We're done with reading the spreadsheet. Now analyze the data.
    # ******
    # Analyze the Schedule cells
    # The first row of the spreadsheet is the list of rooms.
    # Make a list of room names and room column indexes
    roomIndexes: list[int]=[]
    for i in range(0, len(scheduleCells[0])):
        if scheduleCells[0][i] is None:
            break
        if len(scheduleCells[0][i]) > 0:
            roomIndexes.append(i)

    # Get the room names which are in the first row of the scheduleCells tab
    gRoomNames: list[str]=[r.strip() for r in scheduleCells[0]]

    if len(gRoomNames) == 0 or len(roomIndexes) == 0:
        LogError("Room names line is blank.")

    # Start reading ths spreadsheet and building the participants and items databases (dictionaries)
    gSchedules: dict[str, list[ScheduleItem]]=defaultdict(list)  # A dictionary keyed by a person's name containing a ScheduleItem list
    # ScheduleItem is the (time, room, item, moderator) tuples, of an item that that person is on.
    # Note that time and room are redundant and could be pulled out of the Items dictionary
    gItems: dict[str, Item]={}  # A dictionary keyed by item name containing an Item (time, room, people-list, moderator), where people-list is the list of people on the item
    gTimes: list[float]=[]  # A list of times found in the spreadsheet.

    # .......
    # Code to process a set of time and people rows.
    def ProcessRows(timeRow: list[str], peopleRow: Optional[list[str]]) -> None:
        # Get the time from the timerow and add it to gTimes
        time=NumericTime.TextToNumericTime(timeRow[0])
        if time not in gTimes:
            gTimes.append(time)  # We want to allow duplicate time rows, just-in-case

        # Looking at the rest of the row, there may be text in one or more of the room columns
        for roomIndex in roomIndexes:
            roomName=gRoomNames[roomIndex]
            if roomIndex < len(timeRow):  # Trailing empty cells have been truncated, so better check.
                if len(timeRow[roomIndex]) > 0:  # So does the cell itself contain text?
                    # This has to be an item name since it's a cell containing text in a row that starts with a time and in a column that starts with a room
                    itemName=timeRow[roomIndex]

                    # Does a row indexed by peopleRowIndex exist in the spreadsheet? Does it have enough columns? Does it have anything in the correct column?
                    if peopleRow is not None and len(peopleRow) > roomIndex and len(peopleRow[roomIndex]) > 0:
                        # We indicate items which go for an hour, but have some people in one part and some in another using a special notation in the people list.
                        # Robert A. Heinlein, [0.5] John W. Campbell puts RAH on the hour and JWC a half-hour later.
                        # There is much messiness in this.
                        # We look for the [##] in the people list.  If we find it, we divide the people list in half and create two items with separate plists.
                        plistText=peopleRow[roomIndex]
                        r=RegEx.match("(.*)\[([0-9.]*)](.*)", plistText)
                        if r is None:
                            AddItemWithPeople(gItems, gSchedules, time, roomName, itemName, plistText)
                        else:
                            plist1=r.groups()[0].strip()
                            deltaT=r.groups()[1].strip()
                            plist2=r.groups()[2].strip()
                            AddItemWithPeople(gItems, gSchedules, time, roomName, itemName, plist1)
                            newTime=time+float(deltaT)
                            if newTime not in gTimes:
                                gTimes.append(newTime)
                            # This second instance will need to have a distinct item name, so add {#2} to the item name
                            AddItemWithPeople(gItems, gSchedules, newTime, roomName, itemName+" {#2}", plist2)
                    else:  # We have an item with no people on it.
                        AddItemWithoutPeople(gItems, time, roomName, itemName)


    # Now process the schedule row by row
    # When we find a row with data in column 0, we have found a new time.
    # A time row contains items.
    # A time row will normally be followed by a people row containing the participants for those items
    rowIndex: int=1  # We skip the first row which contains room names
    timeRow: Optional[list[str]]=None
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
                LogError("Error reading schedule tab: Row "+str(rowIndex+1)+" is a people row; we were expecting a time row.")     # +1 because the spreadsheet's row-numbering is 1-based
                LogError("   row="+" ".join(row))
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
    count: int=0
    fname=os.path.join(reportsdir, "Diag - precis without items.txt")
    with open(fname, "w") as xml:
        print("Precis without corresponding items:", file=xml)
        for row in precisCells:
            row=[r.strip() for r in row]    # Get rid of leading and trailing blanks
            if len(row[0]) > 0 and len(row[1]) > 0: # If both the item name and the precis exist, store them in the precis table.
                itemname=row[0]
                found=False
                for iname in gItems:
                    if iname == itemname:
                        gItems[itemname].Precis=row[1]
                        found=True
                if not found:
                    count+=1
                    print("   "+itemname, file=xml)
        if count == 0:
            print("    None found", file=xml)

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
    fnameCol=-1
    lnameCol=-1
    emailCol=-1
    responseCol=-1
    fullnameCol=-1
    for i, cell in enumerate(peopleCells[firstNonEmptyRow]):
        cell=cell.lower()
        if cell == "fname":
            fnameCol=i
        if cell == "lname":
            lnameCol=i
        if cell == "full name":
            fullnameCol=i
        if cell == "email":
            emailCol=i
        if cell == "response":
            responseCol=i
    if fnameCol == -1 or lnameCol == -1 or emailCol == -1 or responseCol == -1 or fullnameCol == -1:
        LogError("People tab is missing at least one column label.")
        LogError("    labels="+" ".join(peopleCells[firstNonEmptyRow]))

    @dataclass
    class Person:
        Email: str=""
        RespondedYes: bool=False


    # We'll use the "full name" or, failing that, combine the first and last names to create a full name like is used elsewhere.
    peopleTable: defaultdict[str, Person]=defaultdict(Person)
    for i in range(firstNonEmptyRow+1, len(peopleCells)):
        if len(peopleCells) == 0:   # Skip empty rows
            continue
        row=[r.strip() for r in peopleCells[i]]    # Get rid of leading and trailing blanks in each cell

        fullname=""
        if fullnameCol >= 0 and fullnameCol < len(row):
            fullname=row[fullnameCol]

        if fullname == "":
            fname=""
            if fnameCol >= 0 and fnameCol < len(row):
                fname=row[fnameCol]
            lname=""
            if lnameCol >= 0 and lnameCol < len(row):
                lname=row[lnameCol]
            if len(fname) > 0 and len(lname) > 0:   # Gotta handle Ctein!
                fullname=fname+" "+lname
            elif len(fname) > 0:
                fullname=fname
            elif len(lname) > 0:
                fullname=lname

        if len(fullname.strip()) == 0:
            LogError("Name missing from People tab row #"+str(i+1))
            LogError("    row="+" ".join(peopleCells[i]))

        email=""
        if emailCol >= 0 and emailCol < len(row):
            email=row[emailCol]
        response=""
        if responseCol >= 0 and responseCol < len(row):
            response=row[responseCol]

        if fullname != "":
            peopleTable[fullname]=Person(email, response.lower() == "y")       # Store the email and response as a tuple in the entry indexed by the full name


    #*************************************************************************************************
    #*************************************************************************************************
    # Generate reports
    # The first reports are all error reports or checking reports

    #******
    # Check for people in the schedule who are not in the people tab
    fname=os.path.join( reportsdir, "Diag - People in schedule  but not in People.txt")
    with open(fname, "w") as txt:
        print("People who are scheduled but not in People:", file=txt)
        print("(Note that these may be due to spelling differences, use of initials, etc.)", file=txt)
        count=0
        for personname in gSchedules.keys():
            if personname not in peopleTable.keys():
                count+=1
                print("   "+personname, file=txt)
        if count == 0:
            print("    None found", file=txt)


    #******
    # Check for people in the schedule whose response is not 'y'
    fname=os.path.join( reportsdir, "Diag - People in schedule and in People but whose response is not 'y'.txt")
    with open(fname, "w") as txt:
        print("People who are scheduled and in People but whose response is not 'y':", file=txt)
        count=0
        for personname in gSchedules.keys():
            if personname in peopleTable.keys():
                if peopleTable[personname].RespondedYes:
                    count+=1
                    print(f"   {personname} has a response of {peopleTable[personname].RespondedYes}", file=txt)
        if count == 0:
            print("    None found", file=txt)


    #******
    # Check for people in the schedule whose response is 'y', but who are not scheduled to be on the program
    fname=os.path.join( reportsdir, "Diag - People response is 'y' but who are not scheduled.txt")
    with open(fname, "w") as txt:
        print("People who are scheduled and in People but whose response is 'y' but who are not scheduled:", file=txt)
        count=0
        for personname in peopleTable.keys():
            if peopleTable[personname].RespondedYes:
                found=False
                for item in gSchedules.values():
                    for x in item:
                        if personname == x.PersonName:
                            found=True
                            break
                    if found:
                        break
                if not found:
                    count+=1
                    print(f"   {personname} is not scheduled", file=txt)
        if count == 0:
            print("    None found", file=txt)

    #******
    # Check for people who are scheduled opposite themselves
    fname=os.path.join( reportsdir, "Diag - People scheduled against themselves.txt")
    with open(fname, "w") as txt:
        print("People who are scheduled to be in two places at the same time", file=txt)
        count=0
        for personname in gSchedules.keys():
            pSched=gSchedules[personname] # pSched is a person's schedule, which is a list of (time, room, item) tuples
            if len(pSched) < 2:     # If the persons is only on one item, then there can't be a conflict
                continue
            # Sort pSched by time
            pSched.sort(key=lambda x: x.Time)
            # Look for duplicate times
            prev: ScheduleItem=pSched[0]
            for item in pSched[1:]:
                if item.Time == prev.Time:
                    print(f"{personname}: {NumericTime.NumericToTextDayTime(prev.Time)}: {prev.Room} and also {item.Room}", file=txt)
                    count+=1
                prev=item

        # To make it clear that the test ran, write a message if no conflicts were found.
        if count == 0:
            print("    None found", file=txt)


    #******
    # Now look for similar name pairs
    # First we make up a list of all names that appear in any tab
    names=set()
    names.update(gSchedules.keys())
    names.update(peopleTable.keys())
    similarNames: list[tuple[str, str, float]]=[]
    for p1 in names:
        for p2 in names:
            if p1 < p2:
                rat=difflib.SequenceMatcher(a=p1, b=p2).ratio()
                if rat > .75:
                    similarNames.append((p1, p2, rat))
    similarNames.sort(key=lambda x: x[2], reverse=True)

    fname=os.path.join( reportsdir, "Diag - Disturbingly similar names.txt")
    SafeDelete(fname)
    if len(similarNames) > 0:
        txt=open(fname, "w")
        print("Names that are disturbingly similar:", file=txt)
        count=0
        for s in similarNames:
            print(f"   {s[0]}  &  {s[1]}", file=txt)
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
    fname=os.path.join( reportsdir, "People with items by time.txt")
    SafeDelete(fname)
    with open(fname, "w") as txt:
        for personname in sortedallpartlist:
            print("\n"+personname, file=txt)
            for schedItem in gSchedules[personname]:
                print(f"    {NumericTime.NumericToTextDayTime(schedItem.Time)}: {schedItem.DisplayName} [{schedItem.Room}] {schedItem.ModFlag}", file=txt)

    #*******
    # Print the Items with people by time report
    # Get a list of the program participants (the keys of the  participants dictionary) sorted by the last token in the name (which will usually be the last name)
    fname=os.path.join( reportsdir, "Items with people by time.txt")
    SafeDelete(fname)
    with open(fname, "w") as txt:
        for time in gTimes:
            for room in gRoomNames:
                # Now search for the program item and people list for this slot
                for itemName, item in gItems.items():
                    if item.Time == time and item.Room == room:
                        print(f"{NumericTime.NumericToTextDayTime(time)}, {room}: {itemName}   {item.DisplayPlist()}", file=txt)
                        if item.Precis is not None:
                            print("     "+item.Precis, file=txt)


    #*******
    # Print the program participant's schedule report
    # Get a list of the program participants (the keys of the  participants dictionary) sorted by the last token in the name (which will usually be the last name)
    fname=os.path.join( reportsdir, "Program participant schedules.txt")
    SafeDelete(fname)
    txt=open(fname, "w")
    for personname in sortedallpartlist:
        print("\n\n********************************************", file=txt)
        print(personname, file=txt)
        for schedItem in gSchedules[personname]:
            print(f"\n{NumericTime.NumericToTextDayTime(schedItem.Time)}: {schedItem.DisplayName} [{schedItem.Room}] {schedItem.ModFlag}", file=txt)
            item=gItems[schedItem.ItemName]
            print("Participants: "+item.DisplayPlist(), file=txt)
            if item.Precis is not None:
                print("Precis: "+item.Precis, file=txt)
    txt.close()



    # *******
    # Print the program participant's schedule report
    # Get a list of the program participants (the keys of the  participants dictionary) sorted by the last token in the name (which will usually be the last name)
    fname=os.path.join(reportsdir, "Program participant schedules.xml")
    SafeDelete(fname)
    with open(fname, "w") as xml:
        for personname in sortedallpartlist:
            print(f"<person><full name>{personname}</full name>", file=xml)
            print(f"<email>{peopleTable[personname].Email}</email>", file=xml)
            for schedItem in gSchedules[personname]:
                print(f"<item><title>{NumericTime.NumericToTextDayTime(schedItem.Time)}: {schedItem.DisplayName} [{schedItem.Room}] {schedItem.ModFlag}</title>", file=xml)
                item=gItems[schedItem.ItemName]
                print(f"<participants>{item.DisplayPlist()}</participants>", file=xml)
                if item.Precis is not None:
                    print(f"<precis>{item.Precis}</precis>", file=xml)
                print(f"</item>\n", file=xml)
            print("</person>", file=xml)


    #*******
    # Put out the entire People table in pseudo-XML format
    fname=os.path.join( reportsdir, "Program participants.xml")
    SafeDelete(fname)
    with open(fname, "w") as xml:
        # The first row contains the column headers
        headers=peopleCells[0]
        peopleCells=peopleCells[1:]
        for row in peopleCells:
            xml.writelines(f"<person>")
            for col, cell in enumerate(row):
                xml.writelines(f"<{headers[col]}>{cell}</{headers[col]}>")
            xml.writelines("</person>\n")


    #******
    # Report on the number of people/item
    fname=os.path.join( reportsdir, "Items' people counts.txt")
    SafeDelete(fname)
    txt=open(fname, "w")
    print("List of number of people scheduled on each item\n\n", file=txt)
    for itemname, item in gItems.items():
        print(f"{NumericTime.NumericToTextDayTime(item.Time)} {item.Name}: {len(item.People)}", file=txt)
    txt.close()

    #******
    # Flag items with a suspiciously small number of people on them
    fname=os.path.join( reportsdir, "Diag - Items with unexpectedly low number of participants.txt")
    SafeDelete(fname)
    txt=open(fname, "w")
    print("List of non-readings and KKs with fewer than 3 people on them\n\n", file=txt)
    found=False
    for itemname, item in gItems.items():
        if len(item.People) >= 3:
            continue
        if item.Name.find("Reading") > -1 or item.Name.find("KK") > -1 or item.Name.find("Kaffe") > -1 or item.Name.find("Autograph") > -1:
            continue
        print(f"{NumericTime.NumericToTextDayTime(item.Time)} {item.Name}: {len(item.People)}", file=txt)
        found=True
    if not found:
        print("None found", file=txt)
    txt.close()


    #******
    # Flag items missing a moderator or a precis
    fname=os.path.join( reportsdir, "Diag - Items missing a moderator.txt")
    SafeDelete(fname)
    txt=open(fname, "w")
    print("List of non-readings and KKs with no moderator\n\n", file=txt)
    found=False
    for itemname, item in gItems.items():
        if item.Name.find("Reading") > -1 or item.Name.find("KK") > -1 or item.Name.find("Kaffe") > -1 or item.Name.find("Autograph") > -1:
            continue
        if item.ModName is not None:
            continue
        print(f"{NumericTime.NumericToTextDayTime(item.Time)} {item.Name}: {len(item.People)}", file=txt)
        found=True
    if not found:
        print("None found", file=txt)
    txt.close()

    fname=os.path.join( reportsdir, "Diag - Items missing a precis.txt")
    SafeDelete(fname)
    txt=open(fname, "w")
    print("List of non-readings and KKs with no precis\n\n", file=txt)
    found=False
    for itemname, item in gItems.items():
        if item.Name.find("Reading") > -1 or item.Name.find("KK") > -1 or item.Name.find("Kaffe") > -1 or item.Name.find("Autograph") > -1:
            continue
        if item.Precis is not None and len(item.Precis) > 0:
            continue
        print(f"{NumericTime.NumericToTextDayTime(item.Time)} {item.Name}: {len(item.People)}", file=txt)
        found=True
    if not found:
        print("None found", file=txt)
    txt.close()

    #******
    # Report on the number of items/person
    # Include all people in the people tab, even those with no items
    fname=os.path.join( reportsdir, "Peoples' item counts.txt")
    SafeDelete(fname)
    txt=open(fname, "w")
    print("List of number of items each person is scheduled on\n\n", file=txt)
    for personname in peopleTable:
        if personname in gSchedules.keys():
            print(f"{personname}: {len(gSchedules[personname])}{'' if peopleTable[personname].RespondedYes else ' not confirmed'}", file=txt)
        else:
            if peopleTable[personname].RespondedYes:
                print(personname+": coming, but not scheduled", file=txt)
    txt.close()


    # def Popup(s: str):
    #     return
    #     ctypes.windll.user32.MessageBoxW(0, s, "Your title", 1)
    #
    # Popup("About to create Document()")
    doc=docx.Document()
    # Popup("Document created")
    fname=os.path.join( reportsdir, "Pocket program.txt")
    try:
        # Popup("About to try SafeDelete("+fname+")")
        if not SafeDelete(fname):
            pass
            # Popup("SafeDelete returned False")
    except:
        pass
        # Popup("SafeDelete threw exception")
    try:
        # Popup("About to try Open("+fname+")")
        txt=open(fname, "w")
    except:
        pass
        # Popup("open("+fname+")  threw exception")

    AppendParaToDoc(doc, "Schedule", bold=True, size=24)
    print("Schedule", file=txt)
    for time in gTimes:
        AppendParaToDoc(doc, "")
        AppendParaToDoc(doc, NumericTime.NumericToTextDayTime(time), bold=True)
        print("\n"+NumericTime.NumericToTextDayTime(time), file=txt)
        for room in gRoomNames:
            # Now search for the program item and people list for this slot
            for itemName, item in gItems.items():
                if item.Time == time and item.Room == room:
                    para=doc.add_paragraph()
                    AppendTextToPara(para, room+": ", italic=True, size=12, indent=0.3)
                    AppendTextToPara(para, item.DisplayName, size=12, indent=0.3)
                    print(f"   {room}:  {item.DisplayName}", file=txt)   # Print the room and item name
                    if len(item.People) > 0:            # And the item's people list
                        plist=item.DisplayPlist()
                        AppendParaToDoc(doc, plist, size=12, indent=0.6)
                        print("            "+plist, file=txt)
                    if item.Precis is not None:
                        AppendParaToDoc(doc, item.Precis, italic=True, size=12, indent=0.6)
                        print("            "+item.Precis, file=txt)
    # Popup("About to create Pocket Program.docx")
    fname=os.path.join( reportsdir, "Pocket program.docx")
    doc.save(fname)
    # Popup("Pocket Program.docx has been saved")
    txt.close()


    #******
    # Generate web pages, one for each day.
    currentday=""
    f=None
    for time in gTimes:
        # We generate a separate report for each day
        # The times are sorted in ascending order.
        # We will let the act of the time flipping over to a new day create the new file
        sortday=NumericTime.NumericTimeToNominalDay(time)
        if sortday != currentday:
            # Close the old file, if any
            if f is not None:
                f.write('</font></table>\n')
                # Read and append the footer
                try:
                    with open(PyiResourcePath("control-WebpageFooter.txt")) as f2:
                        f.writelines(f2.readlines())
                except:
                    # wx.App(False)
                    # wx.MessageBox("Can't read 'control-WebpageFooter.txt'")
                    LogError("Can't read 'control-WebpageFooter.txt' (1)")
                f.close()
                f=None
            # And open the new file
            currentday=sortday
            fname=os.path.join( reportsdir, "Schedule - "+sortday+".html")
            SafeDelete(fname)
            f=open(fname, "w")
            try:
                with open(PyiResourcePath("control-WebpageHeader.txt")) as f2:
                    try:
                        f.writelines(f2.readlines())
                    except:
                        LogError("Failure copying 'control-WebpageHeader.txt'")
            except:
                # wx.App(False)
                # wx.MessageBox("Can't read 'control-WebpageHeader.txt'")
                LogError("Can't open 'control-WebpageHeader.txt'")
            f.write("<h2>"+sortday+"</h2>\n")
            f.write('<table border="0" cellspacing="0" cellpadding="2">\n')

        f.write('<tr><td colspan="3">')
        f.write(f'<p class="time">{NumericTime.NumericToTextTime(time)}</p>')
        f.write('</td></tr>\n')
        for room in gRoomNames:
            # Now search for the program item and people list for this slot
            for itemName, item in gItems.items():
                if item.Time == time and item.Room == room:
                    f.write('<tr><td width="40">&nbsp;</td><td colspan="2">')   # Two columns, the first 40 pixes wide and empty
                    f.write(f'<p><span class="room">{room}: </span><span class="item">{item.DisplayName}</span></p>')
                    f.write('</td></tr>')
                    if len(item.People) > 0:            # And the item's people list
                        f.write('<tr><td width="40">&nbsp;</td><td width="40">&nbsp;</td><td width="600">')     # Three columns, the first two 40 pixes wide and empty; the third 600 pixels wide
                        f.write(f'<p><span class="people">{item.DisplayPlist()}</span></p>')
                        f.write('</td></tr>\n')
                    if item.Precis is not None:
                        f.write('<tr><td width="40">&nbsp;</td><td width="40">&nbsp;</td><td width="600">')     # Same
                        f.write(f'<p><span class="precis">{item.Precis}</span></p>')
                        f.write('</td></tr>\n')
    if f is not None:
        # Read and append the footer
        f.write('</table>\n')
        try:
            with open(PyiResourcePath("control-WebpageFooter.txt")) as f2:
                f.writelines(f2.readlines())
        except:
            # wx.App(False)
            # wx.MessageBox("Can't read 'control-WebpageFooter.txt'")
            LogError("Can't read 'control-WebpageFooter.txt' (2)")
        f.close()


    #******
    # Do the room signs.  They'll go in reports/rooms/<name>.docx
    # Create the roomsigns subfolder if none exists
    path=os.path.join( reportsdir, "roomsigns")
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
                    AppendTextToPara(para, NumericTime.NumericToTextDayTime(item.Time)+":  ", bold=True)   # Add the time in bold followed by the item's title
                    AppendTextToPara(para, item.DisplayName)
                    AppendParaToDoc(doc, item.DisplayPlist(), italic=True, indent=0.5)        # Then, on a new line, the people list in italic
        fname=os.path.join(path, room.replace("/", "-")+".docx")
        SafeDelete(fname)
        if inuse:
            doc.save(fname)

    Log(f"Reports generated in directory '{reportsdir}'")
    Log("Done.")
    LogClose()


#*************************************************************************************************
#*************************************************************************************************
# Miscellaneous helper functions

# Read the contents of a spreadsheet tab into
def ReadSheetFromTab(sheet, spreadSheetID, parms: dict[str, str], parmname: str) -> list[list[str]]:

    if parmname not in parms.keys():
        LogError(f"Parameter {parmname} not found in parameters.txt")
        return []

    # Convert the generic name of the tab to the specific name to be used this year
    tabname=parms[parmname]
    cells=[]
    try:
        cells=sheet.values().get(spreadsheetId=spreadSheetID, range=f'{tabname}!A1:Z999').execute().get('values', [])  # Read the whole thing.
    except HttpError as e:
        LogError(f"Can't locate {tabname} tab in spreadsheet. Is the supplied SheetID wrong?")
        exit(999)

    if not cells:
        LogError(f"Can't locate {tabname} tab in spreadsheet")
        raise (ValueError, "No precisCells found")

    return [p for p in cells if len(p) > 0 and "".join(p)[0] != "#"]  # Drop blank lines and lines with a "#" alone in column 1.if not precisCells:



# Generate the name of a person stripped if any "(M)" or "(m)" flags
def RemoveModFlag(s: str) -> str:
    return s.replace("(M)", "").replace("(m)", "").strip()

# Is this person's name flagged as a moderator?
# Check by seeing if RemoveModFlag() does anything
def IsModerator(s: str) -> bool:
    return s != RemoveModFlag(s)

# Delete a file, ignoring any errors
# We do this because of as-yet not understood failures to delete files
def SafeDelete(fn: str) -> bool:
    try:
        os.remove(fn)
    except:
        return False
    return True


#.......
# Add an item with a list of people to the gItems dict, and add the item to each of the persons who are on it
def AddItemWithPeople(gItems: dict[str, Item], gSchedules: dict[str, list[ScheduleItem]], time: float, roomName: str, itemName: str, plistText: str) -> None:

    plist=plistText.split(",")  # Get the people as a list
    plist=[p.strip() for p in plist]  # Remove excess spaces
    plist=[p for p in plist if len(p) > 0]
    modName=""
    peopleList: list[str]=[]
    for person in plist:  # For each person listed on this item
        if IsModerator(person):
            modName=person=RemoveModFlag(person)
        gSchedules[person].append(ScheduleItem(PersonName=person, Time=time, Room=roomName, ItemName=itemName, IsMod=(person == modName)))  # And append a tuple with the time, room, item name, and moderator flag
        peopleList.append(person)
    # And add the item with its list of people to the items table.
    if itemName in gItems:  # If the item's name is already in use, add a uniquifier of room+day/time
        itemName=itemName+"  {"+roomName+" "+NumericTime.NumericToTextDayTime(time)+"}"
    gItems[itemName]=Item(Name=itemName, Time=time, Room=roomName, People=peopleList, ModName=modName)


#.......
# Add an item with a list of people, and add the item to each of the persons
def AddItemWithoutPeople(gItems: dict[str, Item], time: float, roomName: str, itemName: str) -> None:
    if itemName in gItems:  # If the item's name is already in use, add a uniquifier of room+day/time
        itemName=itemName+"  {"+roomName+" "+NumericTime.NumericToTextDayTime(time)+"}"
    gItems[itemName]=Item(Name=itemName, Time=time, Room=roomName)


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



if __name__ == "__main__":
    main()