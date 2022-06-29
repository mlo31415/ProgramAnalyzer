from HelpersPackage import ParmDict

class Person:
    def __init__(self, Fullname: str="", Parms: ParmDict=None):
        self.ListScheduleElement=[]
        self.Fullname=Fullname
        if Parms is None:
            Parms=ParmDict()
        self.Parms: ParmDict=Parms         # This will be a dictionary of *all* columns in the People tab
        pass

    @property
    def RespondedYes(self) -> bool:
        return self.Parms["response", ""].lower() == "y"

    @property
    def Email(self) -> str:
        return self.Parms["email"]

    @property
    def Response(self) -> str:
        return self.Parms["response"]