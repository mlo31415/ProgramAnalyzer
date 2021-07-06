# A class to hold the information for one Participant

from typing import List

from dataclasses import dataclass, field

@dataclass(order=False)
class ScheduleItem:
    PersonName: str=""          # The person's name
    Time: float=None        # A numeric time
    Room: str=""          # The name of a room
    ItemName: str=""      # The name of an item
    IsMod: bool=False       # Is this person the moderator of this item?

    @property
    # Generate the display-name of an item. (Remove any text following the first "{")
    def DisplayName(self):
        loc=self.ItemName.find("{")
        if loc > 0:
            return self.ItemName[:loc-1]
        return self.ItemName

    @property
    def ModFlag(self) -> str:
        if self.IsMod:
            return " (moderator)"
        return ""


