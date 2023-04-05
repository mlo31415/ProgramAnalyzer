# A class to hold the information for one Participant

from dataclasses import dataclass, field

@dataclass(order=False)
class ScheduleElement:
    PersonName: str=""          # The person's name
    Time: float=-1        # A numeric time
    Length: float=1.0       # Length of item in hours
    Room: str=""          # The name of a room
    ItemName: str=""      # The name of an item
    IsMod: bool=False       # Is this person the moderator of this item?
    IsDummy: bool=False     # Is this a dummy item?

    @property
    # Generate the display-name of an item. (Remove any text following the first "{")
    # " Text " --> "Text"
    # " Text {stuff} " --> "Text"
    # "  {stuff} " --> ""
    def DisplayName(self):
        name=self.ItemName.strip()
        loc=name.find("{")
        if loc == -1:   # Curly bracket not found, return the whole thing
            return name
        if loc == 0:    # Everything in the line is after the curley bracket, return empty string
            return ""
        # Return stuff up to the curly bracket
        return name[:loc-1].strip()

    @property
    def ModFlag(self) -> str:
        if self.IsMod:
            return " (moderator)"
        return ""


