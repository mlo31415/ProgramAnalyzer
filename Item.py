from __future__ import annotations

from typing import List
from dataclasses import dataclass, field

from HelpersPackage import ParmDict

# A class to hold the information for one Participant

@dataclass(order=False)
class Item:
    Name: str=""          # The item's name
    Time: float=None        # A numeric time
    Room: str=""          # The name of a room
    People: List[str]=field(default_factory=list)       # A list of the names of people on the item
    ModName: str=""       # The name of the moderator of the item
    Precis: str=""        # The item's precis
    parms: ParmDict=field(default_factory=lambda: ParmDict(CaseInsensitiveCompare=True))    # Parameters associated with this item


    # Generate the display text of a list of people
    def DisplayPlist(self):
        s=""
        for person in self.People:
            s=s+(", " if len(s) > 0 else "")+person+(" (M)" if person == self.ModName else "")
        return s

    @property
    # Generate the display-name of an item. (Remove any text following the first "{")
    def DisplayName(self):
        loc=self.Name.find("{")
        if loc > 0:
            return self.Name[:loc-1]
        return self.Name