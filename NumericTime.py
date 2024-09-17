from typing import Tuple, Any

import re
import math

from HelpersPackage import IsInt, Int0
from Log import LogError, Log

class NumericTime:
    gDayList: list[str]=[]      # List of the day names starting with Day One of the convention schedule.  This must be initialized by a call to
    epsilon=0.001
    startDay=0

    # This takes and of the following:
    #   "Saturday", 13.5
    #   35.5
    #   2, 13.5
    #   "Saturday 13.5"
    #   "Saturday 1:30 pm"
    def __init__(self, day: Any=-1, time: float=-1):
        if day == -1 and time == -1:
            Log("Empty NumericTime class initialized")
            self._day=0
            self._time=0
            return

        # If the day is supplied as a numeric string, turn it into a number
        if isinstance(day, str) and IsInt(day):
            day=int(day)

        if time < -0.5: # No time specified, so interpret the day

            if isinstance(day, str):
                # We have a string for the day and no time specified.  Interpret the day as a full day/time definition
                self.TextToNumericTime(day)
                return
            # So we have a numeric day.  This requires that time be specified, also
            if isinstance(day, float) or isinstance(day, int):
                if day < self.epsilon:
                    self._day=0
                    self._time=0
                    return
                self._day=math.floor((day-self.epsilon)/24)
                self._time=day-24*self._day
                return

            assert False

        if isinstance(day, str):
            # We have a text day and a time defined also.  Interpret the day as the name of a day
            self._day=self.StrToDayNumber(day)
            self._time=math.floor(time+self.epsilon)
            return

        # So day and time are both numeric and supplied
        assert time > -self.epsilon
        self._day=math.floor(day+self.epsilon)
        self._time=math.floor(time+self.epsilon)


    def __eq__(self, other: object) -> bool:
        if isinstance(other, NumericTime):
            return abs(self._day-other._day) < self.epsilon and abs(self._time-other._time) < self.epsilon
        if isinstance(other, float) or isinstance(other, int):
            return self.__eq__(NumericTime(other))
        return NotImplemented

    def __lt__(self, other) -> bool:
        if self._day+self.epsilon < other._day:
            return True
        if other._day+self.epsilon < self._day:
            return False
        return self._time+self.epsilon < other._time

    def __hash__(self):
        return hash(self.Numeric)

    # We only add intervals to a NumericTime --it maes no sense to add Friday, 2pm to Saturday 10am!
    def __add__(self, other):
        if isinstance(other, float) or  isinstance(other, int):
            return NumericTime(self.Numeric+other)
        return NotImplemented

    # If it gets a number, it subtracts that many hours fromt eh NumericTime,  If it gets anothrer NumericTime, it yields the interval between them
    def __sub__(self, other):
        if isinstance(other, float) or  isinstance(other, int):
            return NumericTime(self.Numeric-other)
        if isinstance(other, NumericTime):
            return self.Numeric-other.Numeric
        return NotImplemented

    def __str__(self):
        return f"{self.gDayList[self.Day]} {self.NumericToTextTime()}"

    def __repr__(self):
        return self.__str__

    @classmethod
    def SetStartingDay(cls, day: str):
        daylist=["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]
        if day not in daylist:
            return False
        cls.gDayList=daylist[daylist.index(day):]
        return True

    @property
    def Numeric(self) -> float:
        return 24*self._day+self._time

    @property
    def Bogus(self) -> bool:
        return self.Numeric < self.epsilon


    def StrToDayNumber(self, dstr: str) -> int:
        dstr=dstr.lower()
        daylist=[x.lower() for x in self.gDayList]
        for i, day in enumerate(daylist):
            if day.startswith(dstr):
                return i
        LogError(f"StrToDayNumber(): Can't interpret '{dstr}' as the name of a day")
        assert False


    # Convert a text date string to numeric
    def TextToNumericTime(self, s: str) -> bool:
        s=s.strip()

        # The date string is of the form Day Hour AM/PM or Day Noon
        day=""
        hour=""
        minutes=""
        suffix=""

        m=re.match(r"^([a-z]+|[0-9])\s*([0-9]+)\s*([a-z]+)$", s, re.IGNORECASE)     # <day> <hr> <am/pm/noon/etc>
        if m is not None:
            day=m.groups()[0]
            hour=m.groups()[1]
            suffix=m.groups()[2]
        else:
            m=re.match(r"^([a-z]+|[0-9])\s*([0-9]+):([0-9]+)\s*([a-z]+)$", s, re.IGNORECASE)    # <day> <hr>:<min> <am/pm/noon/etc>
            if m is not None:
                day=m.groups()[0]
                hour=m.groups()[1]
                minutes=m.groups()[2]
                suffix=m.groups()[3]
            else:
                m=re.match(r"^([a-z]+|[0-9])\s*([a-z]+)$", s, re.IGNORECASE)    # <day> <am/pm/noon/etc>
                if m is not None:
                    day=m.groups()[0]
                    suffix=m.groups()[1]
                else:
                    m=re.match(r"^([a-z]+|[0-9])\s*([0-9]+)\s*$", s, re.IGNORECASE)     # <day> <hr>)
                    if m is not None:
                        day=m.groups()[0]
                        hour=m.groups()[1]
                    else:
                        m=re.match(r"^([a-z]+|[0-9])\s*([0-9]+):([0-9]+)\s*$", s, re.IGNORECASE)  # <day> <hr>:<min>
                        if m is not None:
                            day=m.groups()[0]
                            hour=m.groups()[1]
                            minutes=m.groups()[2]
                        else:
                            m=re.match(r"^([a-z]+|[0-9])\s*([0-9]+).([0-9]+)\s*$", s, re.IGNORECASE)  # <day> <hr>.<fraction>
                            if m is not None:
                                day=m.groups()[0]
                                hour=m.groups()[1]
                                minutes=60*float("."+m.groups()[2])
                            else:
                                LogError("Can't interpret time: '"+s+"'")
                                return False

        d=0
        if IsInt(day):
            d=Int0(day)
        else:
            d=self.StrToDayNumber(day)

        h=0
        if hour != "":
            h=int(hour)
        if minutes != "":
            h=h+int(minutes)/60
        if suffix.lower() == "pm":
            h=h+12
        elif suffix.lower() == "am" and int(hour) == 12:
            h=h-12  # Special case of 12:30 am being 30 minutes into the day
        elif suffix.lower() == "noon":
            h=12
        elif suffix.lower() == "midnight":
            h=24

        #print("'"+s+"'  --> day="+day+"  hour="+hour+"  minutes="+minutes+"  suffix="+suffix+"   --> d="+str(d)+"  h="+str(h)+"  24*d+h="+(str(24*d+h))+"  --> "+NumericToTextDayTime(24*d+h))
        self._day=d
        self._time=h
        return True


    @property
    def Day(self) -> int:
        return self._day

    @property
    def Hour(self) -> float:
        d, h, t, ispm=self.DayHourMinute
        h+=t
        if ispm:
            h+=12
        return h

    @property
    def DayHourMinute(self) -> Tuple[int, int, float, bool]:
        try:
            t=self._time
        except AttributeError:
            pass

        isPM=t>12           # AM or PM?
        if isPM:
            t=t-12
        h=math.floor(t)     # Get the hour
        t=t-h               # What's left is the fractional hour
        return self._day, h, t, isPM


    def NumericToTextTime(self) -> str:
        d, h, m, isPM=self.DayHourMinute

        if h == 12 and m == 0:         # Handle noon and midnight specially
            if isPM:
                return "Midnight"
            else:
                return "Noon"

        if h == 0 and m != 0:
            numerictime=f"12:{math.floor(60*m):02}"     # Handle the special case of times after noon but before 1
        else:
            numerictime=f"{h}{'' if m == 0 else f':{math.floor(60*m):02}'}"

        return numerictime + (" pm" if isPM else " am")


    # Return the name of the day corresponding to a numeric time
    @property
    def DayString(self) -> str:
        return self.gDayList[int(self.Day)]

    # We sort days based on one day ending and the next beginning at 4am -- this puts late-night items with the previous day
    # Note that the return value is used for sorting, but not for dae display
    @property
    def NominalDayString(self) -> str:
        if self.Numeric > 4:
            return (self-4).DayString
        return self.gDayList[0]     # This is wrong, but what can we do?