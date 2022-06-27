class Person:
    def __init__(self, Email: str="", Response: str=""):
        self.ListScheduleElement=[]
        self.Email: str=Email
        self.Response: str=Response
        pass

    @property
    def RespondedYes(self) -> bool:
        return self.Response.lower() == "y"

