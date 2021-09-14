from typing import Tuple

import re
import math

from Log import LogError


gDayList=["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]

# Convert a text date string to numeric
def TextToNumericTime(s: str) -> int:
    # The date string is of the form Day Hour AM/PM or Day Noon
    day=""
    hour=""
    minutes=""
    suffix=""

    m=re.match(r"^([A-Za-z]+)\s*([0-9]+)\s*([A-Za-z]+)$", s)     # <day> <hr> <am/pm/noon/etc>
    if m is not None:
        day=m.groups()[0]
        hour=m.groups()[1]
        suffix=m.groups()[2]
    else:
        m=re.match(r"^([A-Za-z]+)\s*([0-9]+):([0-9]+)\s*([A-Za-z]+)$", s)    # <day> <hr>:<min> <am/pm/noon/etc>
        if m is not None:
            day=m.groups()[0]
            hour=m.groups()[1]
            minutes=m.groups()[2]
            suffix=m.groups()[3]
        else:
            m=re.match(r"^([A-Za-z]+)\s*([A-Za-z]+)$", s)    # <day> <am/pm/noon/etc>
            if m is not None:
                day=m.groups()[0]
                suffix=m.groups()[1]
            else:
                LogError("Can't interpret time: '"+s+"'")

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


def DayNumber(t: float) -> int:
    return math.floor((t-.01)/24)  # Compute the day number. The "-.01" is to force midnight into the preceding day rather than the following day

# Convert a numeric daytime to text
# The input time is a floating point number of hours since the start of the 1st day of the convention
def NumericToTextDayTime(t: float) -> str:
    return f"{gDayList[DayNumber(t)]} {NumericToTextTime(t)}"


def NumericTimeToDayHourMinute(t: float) -> Tuple[int, int, float, bool]:
    d=DayNumber(t)
    t=t-24*d
    isPM=t>12           # AM or PM?
    if isPM:
        t=t-12
    h=math.floor(t)     # Get the hour
    t=t-h               # What's left is the fractional hour
    return d, h, t, isPM


def NumericToTextTime(f: float) -> str:
    d, h, m, isPM=NumericTimeToDayHourMinute(f)

    if h == 12:         # Handle noon and midnight specially
        if isPM:
            return "Midnight"
        else:
            return "Noon"

    if h == 0 and m != 0:
        numerictime="12:"+str(math.floor(60*m))     # Handle the special case of times after noon but before 1
    else:
        numerictime=str(h) + ("" if m == 0 else ":" + str(math.floor(60*m)))

    return numerictime + (" pm" if isPM else " am")


# Return the name of the day corresponding to a numeric time
def NumericTimeToDayString(f: float) -> str:
    d, _, _, _=NumericTimeToDayHourMinute(f)
    return gDayList[int(d)]

# We sort days based on one day ending and the next beginning at 4am -- this puts late-night items with the previous day
# Note that the return value is used for sorting, but not for dae display
def NumericTimeToNominalDay(f: float) -> str:
    return NumericTimeToDayString(f-4)