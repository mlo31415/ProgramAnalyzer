# A class to hold the information for one Participant

from dataclasses import dataclass, field

@dataclass(order=False)
class ScheduleItem:
    PersonName: str=None          # The person's name
    Time: float=None        # A numeric time
    Room: str=None          # The name of a room
    ItemName: str=None      # The name of an item
    IsMod: bool=False       # Is this person the moderator of this item?


    def __init__(self, PersonName:str=None, Time:float=None, Room:str=None, ItemName:str=None, IsMod:bool=False):
        self.PersonName=PersonName
        self.Time=Time
        self.Room=Room
        self.ItemName=ItemName
        self.IsMod=IsMod

    @property
    # Generate the display-name of an item. (Remove any text following the first "{")
    def DisplayName(self):
        loc=self.ItemName.find("{")
        if loc > 0:
            return self.ItemName[:loc-1]
        return self.ItemName

