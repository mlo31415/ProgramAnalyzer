from HelpersPackage import ParmDict, YesNoMaybe

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