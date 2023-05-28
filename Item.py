from __future__ import annotations

from HelpersPackage import ParmDict, SearchAndReplace
from NumericTime import NumericTime
from Log import Log

# A class to hold the information for one Item
class Item:
    def __init__(self, ItemText: str="", Time: NumericTime=None, Length: float=1.0, Room: str="", People: list[str]=None, ModName: str="", Precis: str="", Parms: ParmDict=None):
        self.Time: NumericTime=Time
        self.Length: float=Length
        self.Room: str=Room
        if People is None:
            People=[]
        self.People: list[str]=People   # List of keys of People on this item
        self.ModName: str=ModName
        self.Precis: str=Precis
        if Parms is None:
            Parms=ParmDict(CaseInsensitiveCompare=True)
        self.Parms: ParmDict=Parms
        self.ItemText=ItemText  # This must be last as it relies on the rest of the object having been initialized
        self.IsContinuation: bool=False
        if "{cont}" in ItemText:
            self.IsContinuation=True


    @property
    def ItemText(self) -> str:
        return self._ItemText
    @ItemText.setter
    # We take the entire item cell contents and parse it into its pieces.
    def ItemText(self, val: str):
        # Save the whole item text in _ItemText
        self._ItemText=val
        # Initialize the rest to empty strings
        self._Name: str=""
        self._Comment=""

        # Strip off the comment.  This must follow the last "}" so we can say things like '{#2}'
        if "#" in val:
            loc=val.find("#")
            loc2=val.rfind("}")
            if loc > loc2:
                self._Comment=val[loc-1:]
                val=val[:loc-1].strip()
                if len(val) == 0:
                    return

        # Look for keywords, remove them and save them
        lst, val=SearchAndReplace("(<.*?>)", val, "")
        val=val.strip()
        for l in lst:
            l=l.strip("<>").strip()
            if ":" in l:
                loc=l.find(":")
                self.Parms[l[:loc]]=l[loc+1:].strip()
            else:
                self.Parms[l]="True"

        # And save the name
        self._Name=val.strip()

    @property
    def Name(self) -> str:
        return self._Name


    # Generate the display text of a list of people
    def DisplayPlist(self):
        s=""
        for person in self.People:
            s=s+(", " if len(s) > 0 else "")+person+(" (M)" if person == self.ModName else "")
        return s


    @property
    # Generate the display-name of an item. (Remove any text following the first "{")
    # " Text " --> "Text"
    # " Text {stuff} " --> "Text"
    # "  {stuff} " --> ""
    def DisplayName(self):
        name=self.Name.strip()
        loc=name.find("{")
        if loc == -1:   # Curly bracket not found, return the whole thing
            return name
        if loc == 0:    # Everything in the line is after the curley bracket, return empty string
            return ""
        # Return stuff up to the curly bracket
        return name[:loc-1].strip()
