import re as RegEx

from HelpersPackage import ParmDict, YesNoMaybe, Int0
from NumericTime import TextToNumericTime

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
    def Avoid(self) -> list[tuple[float, float]]:
        if "avoid" not in self.Parms:
            return []
        return ParseAvoid(self.Parms["avoid"])

# Parse the Avoid column for a person into times to be avoided.
def ParseAvoid(avstring: str) -> list[tuple[float, float]]:
    # The contents are a list of comma-separated times or time-ranges.  First create the list of individual items and remove excess spaces.
    avstrl=[x.strip() for x in avstring.split(",")]

    # Individual avoid strings can be of the following forms:
    # All are case-insensitive. Times are numeric int or float, 24-hour clock
    # Arrive: [day] [time]      (If day is missing, Friday is assumed)
    # [Leave, Depart]: [day] [time]      (If day is missing, Sunday is assumed)
    # [Day]: float-float | dinner | evening
    out: list[tuple[float, float]]=[]   # A list of start-end tuples
    for avs in avstrl:
        avl=[x.strip().lower() for x in avs.split(" ")]
        assert len(avl) > 0
        command=avl[0]
        avl=avl[1:]
        match command:
            case "arrive":
                # [day] time
                day="fri"
                time=""
                if len(avl) > 1:
                    day=avl[0]
                    time=avl[1]
                else:
                    time=avl[0]
                out.append((0, TextToNumericTime(day+" "+time)))

            case "leave" | "depart":

                # [day] time
                day="sun"
                time=""
                if len(avl) > 1:
                    day=avl[0]
                    time=avl[1]
                else:
                    time=avl[0]
                out.append((TextToNumericTime(day+" "+time), 999))

            case "fri" | "friday":
                # [time-time] | "dinner" | "evening"
                ret=ProcessTimeRange(avl, "fri")
                if ret is not None:
                    out.append(ret)

            case "sat" | "saturday":
                # [time-time] | "dinner" | "evening"
                ret=ProcessTimeRange(avl, "sat")
                if ret is not None:
                    out.append(ret)

            case "sun" | "sunday":
                # [time-time] | "dinner" | "evening"
                ret=ProcessTimeRange(avl, "sun")
                if ret is not None:
                    out.append(ret)

    return out



def ProcessTimeRange(avl: list[str], day: str="") -> tuple[float, float] | None:
    range=()
    if avl[0] == "dinner":
        range=(18, 20)
    elif avl[0] == "evening":
        range=(20, 24)
    else:
        # We probably have a number range (nn-nn)
        m=RegEx.match("([0-9.:]+)-([0-9.:]+)$", avl[0])
        if m is not None:
            range=(float(m.groups()[0]), float(m.groups()[1]))
    if len(range) == 0:
        return ()
    return (TextToNumericTime(f"{day} {range[0]}"), TextToNumericTime(f"{day} {range[1]}"))