# A class to hold the information for one Participant

from dataclasses import dataclass, field

@dataclass(order=False)
class Item:
    Name: str=None          # The item's name
    Time: float=None        # A numeric time
    Room: str=None          # The name of a room
    People: list=None       # A list of the names of people on the item
    ModName: str=None       # The name of the moderator of the item
    Precis: str=None        # The item's precis


    def __init__(self, Name:str=None, Time:float=None, Room:str=None, People:list=None, ModName:str=None, Precis:str=None):
        self.Name=Name
        self.Time=Time
        self.Room=Room
        self.People=People
        self.ModName=ModName
        self.Precis=Precis

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