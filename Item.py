from __future__ import annotations

from HelpersPackage import ParmDict, SearchAndReplace

# A class to hold the information for one Participant

class Item:
    def __init__(self, ItemText: str="", Time: float=0.0, Room: str="", People: list[str]=None, ModName: str="", Precis: str="", Parms: ParmDict=None):
        self.Time: float=Time
        self.Room: str=Room
        if People is None:
            People=[]
        self.People: list[str]=People
        self.ModName: str=ModName
        self.Precis: str=Precis
        if Parms is None:
            Parms=ParmDict(CaseInsensitiveCompare=True)
        self.Parms: ParmDict=Parms
        self.ItemText=ItemText  # This must be last as it relies on the rest of the object having been initialized


    @property
    def ItemText(self) -> str:
        return self._ItemText
    @ItemText.setter
    # We take the entire item cell contents and parse it into its pieces.
    def ItemText(self, val: str):
        # Save the whole item text in _ItenText
        self._ItemText=val
        # Initialize the rest to empty strings
        self._Name: str=""
        self._ItemText: str=""
        self._Comment=""

        # Strip off the comment
        if "#" in val:
            loc=val.find("#")
            self._Comment=val[loc:]
            val=val[:loc].strip()
            if len(val) == 0:
                return

        # Look for keywords, remove them and save them
        lst, val=SearchAndReplace("(<.*?>)", val, "")
        val=val.strip()
        for l in lst:
            l=l.strip("<>").strip()
            if ":" in l:
                loc=l.find(":")
                self.Parms[l[:loc]]=l[loc:]
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