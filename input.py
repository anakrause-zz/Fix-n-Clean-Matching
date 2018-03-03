import pandas as pd

# imports, parses the file and deletes useless columns

# volunteer file

xlsx = pd.ExcelFile('volunteers.xlsx')
df = xlsx.parse(0)
df = df.drop(df.columns[[3, 6, 9, 12]], axis=1)

# members file

members = pd.ExcelFile('members.xlsx')
thingy = members.parse(0)
thingy = thingy.drop(thingy.columns[[7,8,9,10]], axis=1)

#converts pandas dataframe to matrix
np_df_vol = df.as_matrix()
np_df_mem = thingy.as_matrix()

groups = []
members = []

class Group(object):
    def __init__(self, time, name1, email1, name2, email2, name3, email3, name4, email4, name5, email5, av1, av2, if21, netid, info):
        self.time = time
        self.name1 = name1
        self.email1 = email1
        self.name2 = name2
        self.email2 = email2
        self.name3 = name3
        self.email3 = email3
        self.name4 = name4
        self.email4 = email4
        self.name5 = name5
        self.email5 = email5
        self.av1 = av1
        self.av2 = av2
        self.av1num = 0
        self.av2num = 0
        self.if21 = if21
        self.netid = netid
        self.info = info

        self.av1 = str(self.av1)
        self.av2 = str(self.av2)

        def gettime (varin):
            if "Saturday" in varin:
                if "9:00-12:00" in varin:
                    return 1
                elif "1:00-4:00" in varin:
                    return 2
            elif "Sunday" in varin:
                if "9:00-12:00" in varin:
                    return 3
                elif "1:00-4:00" in varin:
                    return 4
            return 0

        self.av1num = gettime(self.av1)
        self.av2num = gettime(self.av2)


    def returngroupinfo(self):
        return (self.time, self.name1, self.email1, self.name2, self.email2, self.name3, self.email3, self.name4, self.email4, self.name5, self.email5, self.av1, self.av2, self.av1num, self.av2num, self.if21, self.netid, self.info)

class Member(object):
    def __init__(self, name, phone, email, methcontact, ifcontacted, ifconfirm, timeslot, task, address, info):
        self.name = name
        self.phone = phone
        self.email = email
        self.methcontact = methcontact
        self.ifcontacted = ifcontacted
        self.ifconfirm = ifconfirm
        self.timeslot = timeslot
        self.numtimeslot = 0
        self.task = task
        self.address = address
        self.info = info
        self.flagged = False
        self.volunteergroup = None

        self.timeslot = str(self.timeslot)
        if "Saturday" in self.timeslot:
            if "Morning" in self.timeslot:
                self.numtimeslot = 1
            elif "Afternoon" in self.timeslot:
                self.numtimeslot = 2
        elif "Sunday" in self.timeslot:
            if "Morning" in self.timeslot:
                self.numtimeslot = 3
            elif "Afternoon" in self.timeslot:
                self.numtimeslot = 4
        else:
            self.flagged = True


    def returnmeminfo(self):
        return (self.name, self.phone, self.email, self.methcontact, self.ifcontacted, self.ifconfirm, self.timeslot, self.numtimeslot, self.task, self.address, self.info, self.flagged)



for g in np_df_vol:
    gr = Group(g[0], g[1], g[2], g[3], g[4], g[5], g[6], g[7], g[8], g[9], g[10], g[11], g[12], g[13], g[14], g[15])
    groups.append(gr)

for m in np_df_mem:
    mem = Member(m[0], m[1], m[2], m[3], m[4], m[5], m[6], m[7], m[8], m[9])
    members.append(mem)


print (groups[0].returngroupinfo())
#print (members[5].returnmeminfo())

