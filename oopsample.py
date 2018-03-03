class User(object):
    def __init__(self, firstname, lastname, email, mobile, homeform):
        self.firstname = firstname
        self.lastname = lastname
        self.email = email
        self.mobile = mobile
        self.homeform = homeform
    def __init__(self, lastname, email, mobile, homeform):
        self.firstname = "hanna"
        self.lastname = lastname
        self.email = email
        self.mobile = mobile
        self.homeform = homeform
    def getinfo(self):
        return [self.firstname, self.lastname, self.email, self.mobile, self.homeform]


class AddressBook():
    def __init__(self):
        self.book = []
    def addUser(self, user):
        self.book.append(user)
    def printAllUsers(self):
        for users in self.book:
            print (users.getinfo())
    def numberOfUsers(self):
        print (len(self.book))

hana = User( "Gill", "hana.gill89@gmail.com", "6473092809", "Ghost")
ana = User("Anastasia", "Krause", "anastasiavkrause@gmail.com", "6478932757", "Ms.Frensch")

novel = AddressBook()
novel.addUser(ana)
novel.addUser(hana)
novel.printAllUsers()
novel.numberOfUsers()
