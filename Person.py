import re as RegEx

from HelpersPackage import ParmDict, YesNoMaybe
from NumericTime import NumericTime


# ======================================================
class Avoidment:
    def __init__(self, start: NumericTime, end: NumericTime, desc: str):
        self.Start: NumericTime=start
        self.End: NumericTime=end
        self.Description=desc

    def __str__(self) -> str:
        return self.Description

    def Pretty(self) -> str:
        out=""
        avs=[(self.Start, self.End)]    # This will be a list of NumericTime tuples
        if self.Start.Day != self.End.Day:  # If the avoidance spans more than one day, break it into single-day avoidances
            avs=[(self.Start, NumericTime(f"{self.Start.Day} 12:59 pm"))]
            for day in range(self.Start.Day+1, self.End.Day+1):
                if day == self.End.Day:
                    avs.append((NumericTime(f"{day} 12:01 am"), self.End))
                else:
                    avs.append((NumericTime(f"{day} 12:01 am"), NumericTime(f"{day} 12:59 pm")))

        for tpl in avs:
            ntstart=tpl[0]
            ntend=tpl[1]

            # Check for the whole day
            if ntstart.Hour < 0.05 and ntend.Hour > 23.95:
                if out != "":
                    out+="; "
                out+=ntstart.DayString+ "(all day)"
                continue

            # OK, we know this is all in one day but is not the whole day.  Handle the cases where we start at midnight or end at midnight
            if ntstart.Hour < 0.05:
                if out != "":
                    out+="; "
                out+=f"--{ntend}"
                continue
            if ntend.Hour > 23.95:
                if out != "":
                    out+="; "
                out+=f"{ntstart} --"
                continue
            if out != "":
                out+="; "
            out+= f"{ntstart} -- {ntend}"

        return out

    @property
    def Duration(self) -> float:
        return self.Start-self.End


# ======================================================
class Person:
    def __init__(self, Fullname: str="", Parms: ParmDict=None):
        self.ListScheduleElement=[]
        self.Fullname=Fullname

        # Parms is a dictionary of *all* columns in the People tab.  If it is not supplied, just create an empty ParmDict()
        if Parms is None:
            Parms=ParmDict()
        self.Parms: ParmDict=Parms
        pass

    @property
    def RespondedYes(self) -> bool:
        return "yes" == YesNoMaybe(self.Parms["response", ""])

    @property
    def Email(self) -> str:
        return self.Parms["email"]

    @property
    def Response(self) -> str:
        return YesNoMaybe(self.Parms["response"])

    @property
    def Avoid(self) -> list[Avoidment]:
        if "avoid" not in self.Parms:
            return []
        return ParseAvoid(self.Parms["avoid"])


# Parse the Avoid column for a person into times to be avoided.
def ParseAvoid(avstring: str) -> list[Avoidment]:
    # The contents are a list of comma-separated times or time-ranges.  First create the list of individual items and remove excess spaces.
    avstrl=[x.strip() for x in avstring.split(",")]

    # Individual avoid strings can be of the following forms:
    # All are case-insensitive. Times are numeric int or float, 24-hour clock
    # Arrive: [day] [time]      (If day is missing, Friday is assumed)
    # [Leave, Depart]: [day] [time]      (If day is missing, Sunday is assumed)
    # [Day]: float-float | dinner | evening
    out: list[Avoidment]=[]   # A list of start-end tuples
    for avs in avstrl:
        avl=[x.strip().lower() for x in avs.split(" ")]
        assert len(avl) > 0
        command=avl[0]
        avl=avl[1:]
        match command:
            case "arrive":
                # [day] time
                day="fri"       #TODO: This and other day references really ought to depend on the starting and ending days of the convention
                time=""
                if len(avl) > 1:
                    day=avl[0]
                    time=avl[1]
                else:
                    time=avl[0]
                if day.startswith("sun"):    # If the arrival day is Sunday, both Friday and Saturday are also excluded
                    out.append(Avoidment(NumericTime("Saturday 12:01 am"), NumericTime("Saturday 11:59 pm"), avs))
                if day.startswith("sat") or day.startswith("sun"):
                    out.append(Avoidment(NumericTime("Friday 12:01 am"), NumericTime("Friday 11:59 pm"), avs))
                out.append(Avoidment(NumericTime(day+" 12:01 am"), NumericTime(day+" "+time), avs))

            case "leave" | "depart":

                # [day] time
                day="sun"
                time=""
                if len(avl) > 1:
                    day=avl[0]
                    time=avl[1]
                else:
                    time=avl[0]
                if day == "sat":
                    out.append(Avoidment(NumericTime("Friday 12:01 am"), NumericTime("Friday 11:59 pm"), "Friday all day"))
                if day == "fri":
                    out.append(Avoidment(NumericTime("Saturday 12:01 am"), NumericTime("Saturday 11:59 pm"), "Saturday all day"))
                out.append(Avoidment(NumericTime(day+" "+time), NumericTime(day+ " 11:59pm"), avs))

            case "fri" | "friday":
                # [time-time] | "dinner" | "evening" | "all day"
                ret=ProcessTimeRange(avl, "fri")
                if ret is None:
                    continue
                ret.Description=avs
                out.append(ret)

            case "sat" | "saturday":
                # [time-time] | "dinner" | "evening" | "all day"
                ret=ProcessTimeRange(avl, "sat")
                if ret is None:
                    continue
                ret.Description=avs
                out.append(ret)

            case "sun" | "sunday":
                # [time-time] | "dinner" | "evening" | "all day"
                ret=ProcessTimeRange(avl, "sun")
                if ret is None:
                    continue
                ret.Description=avs
                out.append(ret)

            case "daily" | "every" | "all":
                for day in ["thu", "fri", "sat", "sun", "mon", "tue", "wed"]:   # A bit of a kludge, but we don't know the actual con days this deep in Person
                    # [time-time] | "dinner" | "evening"
                    ret=ProcessTimeRange(avl, day)
                    if ret is None:
                        continue
                    ret.Description=avs
                    out.append(ret) #
    return out  #Test of push





def ProcessTimeRange(avl: list[str], day: str="") -> Avoidment | None:
    range=()
    if avl[0] == "dinner":
        range=(18, 20)
    elif avl[0] == "evening":
        range=(20, 24)
    elif avl[0] == "all" and avl[1] == "day":
        range=(0.02, 23.98)
    else:
        # We probably have a number range (nn-nn)
        m=RegEx.match("([0-9.:]+)-([0-9.:]+)$", avl[0])
        if m is not None:
            range=(float(m.groups()[0]), float(m.groups()[1]))
    if len(range) == 0:
        return None
    return Avoidment(NumericTime(f"{day} {range[0]}"), NumericTime(f"{day} {range[1]}"), "")

