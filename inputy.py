import pandas as pd
import numpy as np
import math

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

#create structure for members
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
        self.av1num = -1
        self.av2num = -1
        self.avfinal = 0
        self.if21 = if21
        self.netid = netid
        self.info = info
        self.gsize = -1

        self.av1 = str(self.av1)
        self.av2 = str(self.av2)

        # assigns numeric value of 1-4 based on text availability
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
            return -1

        self.av1num = gettime(self.av1)
        self.av2num = gettime(self.av2)

        #fixes people who put down same availability twice
        if (self.av2num == self.av1num):
            self.av2num = -1

        #creates unique pair for each combination of availabilities
        self.avfinal = self.av1num * self.av2num

        if (self.avfinal == -1):
            self.avfinal = 0
        elif (self.avfinal == -2):
            self.avfinal = 1
        elif (self.avfinal == -3):
            self.avfinal = 2
        elif (self.avfinal == -4):
            self.avfinal = 3
        elif (self.avfinal == 2):
            self.avfinal = 4
        elif (self.avfinal == 3):
            self.avfinal = 5
        elif (self.avfinal == 4):
            self.avfinal = 6
        elif (self.avfinal == 6):
            self.avfinal = 7
        elif (self.avfinal == 12):
            self.avfinal = 9

        #assigns size value of each group based on # of 'nan' values
        if (type(self.name5) == float):
            if (type(self.name4) == float):
                if (type(self.name3) == float):
                    if (type(self.name2) == float):
                        self.gsize = 1
                    else:
                        self.gsize = 2
                else:
                    self.gsize = 3
            else:
                self.gsize = 4
        else:
            self.gsize = 5


    def returngroupinfo(self):
        return (self.time, self.name1, self.email1, self.name2, self.email2, self.name3, self.email3, self.name4, self.email4, self.name5, self.email5, self.av1, self.av2, self.av1num, self.av2num, self.avfinal, self.if21, self.netid, self.info, self.gsize)

    def returngroupsize(self):
        return(self.gsize)

    def returnavfinal(self):
        return(self.avfinal)



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

        # assign numeric timeslot based on availability
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

    def returnnumtimeslot(self):
        return self.numtimeslot


def addempty(arr):
    for x in range(10):
        arr.append([])

groupsfive = []
groupsfour = []

groupsthree = []
groupstwo = []
groupsone = []

# adds 10 empty arrays to the array
addempty(groupsfive)
addempty(groupsfour)
addempty(groupsthree)
addempty(groupstwo)
addempty(groupsone)

# iterates over all group arrays from original matrix (each g is a group array)
# gr is a group object that is created with array input g[0], g[1], g[...]
# append certain objects to certain arrays based on groupsize (gsize) and avfinal value
#print (np_df_vol[5])
for g in np_df_vol:
    gr = Group(g[0], g[1], g[2], g[3], g[4], g[5], g[6], g[7], g[8], g[9], g[10], g[11], g[12], g[13], g[14], g[15])
    if gr.returngroupsize() == 5:
        groupsfive[gr.returnavfinal()].append(gr)
    elif gr.returngroupsize() == 4:
        groupsfour[gr.returnavfinal()].append(gr)
    elif gr.returngroupsize() == 3:
        groupsthree[gr.returnavfinal()].append(gr)
    elif gr.returngroupsize() == 2:
        groupstwo[gr.returnavfinal()].append(gr)
    else:
        groupsone[gr.returnavfinal()].append(gr)

#same thing but for members
for m in np_df_mem:
    mem = Member(m[0], m[1], m[2], m[3], m[4], m[5], m[6], m[7], m[8], m[9])
    members.append(mem)

#for x in range(len(groupsfive[5])):
#    print (groupsfive[5][x].returngroupinfo())

#for x in range(5):
    #print (groupsfour[x].returngroupinfo())
#print (members[5].returnmeminfo())

def checknan(teststring):
    if "nan" in teststring:
        return ""
    else:
        return teststring
# test case for 2 and 2
#thefirst = groupstwo[4][0]
#thesecond = groupstwo[4][1]

#test case for 3 and 1
thefirst = groupstwo[4][0]
thesecond = groupsone[4][0]
thethird = groupsone[4][1]
thefourth = groupsone[4][2]
#print (thefirst.returngroupinfo())
#print (thesecond.returngroupinfo())
#print (thethird.returngroupinfo())
#print (thefourth.returngroupinfo())

#                   1       1       2           1
def combineGroups(first, second, third = 0, fourth = 0):
    one = 0
    two = 0
    three = 0
    four = 0
    if first.returngroupsize() < second.returngroupsize():
        two = first
        one = second
    else:
        one = first
        two = second

    if third != 0:
        if first.returngroupsize() == 2:
            one = first
            two = second
            three = third
            four = fourth
        elif second.returngroupsize() == 2:
            one = second
            two = first
            three = third
            four = fourth
        elif third.returngroupsize() == 2:
            one = third
            two = first
            three = second
            four = fourth
        elif fourth.returngroupsize() == 2:
            one = fourth
            two = first
            three = second
            four = third

    ntimestamp = one.returngroupinfo()[0]
    nvol1name = one.returngroupinfo()[1]
    nvol1email = one.returngroupinfo()[2]
    nvol2name = one.returngroupinfo()[3]
    nvol2email = one.returngroupinfo()[4]

    ## COMBINE 3 and 2
    if one.returngroupsize() == 3 and two.returngroupsize() == 2:
        nvol3name = one.returngroupinfo()[5]
        nvol3email = one.returngroupinfo()[6]

        nvol4name = two.returngroupinfo()[1]
        nvol4email = two.returngroupinfo()[2]
        nvol5name = two.returngroupinfo()[3]
        nvol5email = two.returngroupinfo()[4]

    # COMBINE 2 and 2 OR 3 and 1
    if (one.returngroupsize() == 2 and two.returngroupsize() == 2) or (one.returngroupsize() == 3 and two.returngroupsize() == 1):
        if (one.returngroupsize() == 2 and two.returngroupsize() == 2):
            nvol3name = two.returngroupinfo()[1]
            nvol3email = two.returngroupinfo()[2]
            nvol4name = two.returngroupinfo()[3]
            nvol4email = two.returngroupinfo()[4]
        elif (one.returngroupsize() == 3 and two.returngroupsize() == 1):
            nvol3name = one.returngroupinfo()[5]
            nvol3email = one.returngroupinfo()[6]
            nvol4name = two.returngroupinfo()[1]
            nvol4email = two.returngroupinfo()[2]

        nvol5name = float('nan')
        nvol5email = float('nan')

    if (one.returngroupsize() == 2 and two.returngroupsize() == 1 and three.returngroupsize() == 1):
        nvol3name = two.returngroupinfo()[1]
        nvol3email = two.returngroupinfo()[2]
        nvol4name = three.returngroupinfo()[1]
        nvol4email = three.returngroupinfo()[2]

        if four == 0:
            nvol5name = float('nan')
            nvol5email = float('nan')
        else:
            nvol5name = four.returngroupinfo()[1]
            nvol5email = four.returngroupinfo()[2]

    ## DONT CHANGE THIS
    ntime1 = one.returngroupinfo()[11]
    ntime2 = one.returngroupinfo()[12]

    oinfoarr = []
    for x in range (-4, -1, 1):
        oinfo1 = str(one.returngroupinfo()[x])
        oinfo2 = str(two.returngroupinfo()[x])

        oinfoarr.append(oinfo1)
        oinfoarr.append(oinfo2)

    for y in oinfoarr:
        y = checknan(y)

    ninfo1 = oinfoarr[0] + " " + oinfoarr[1]
    ninfo2 = oinfoarr[2] + " " + oinfoarr[3]
    ninfo3 = oinfoarr[4] + " " + oinfoarr[5]

    gr = Group(ntimestamp, nvol1name, nvol1email, nvol2name, nvol2email, nvol3name, nvol3email, nvol4name, nvol4email, nvol5name, nvol5email, ntime1, ntime2, ninfo1, ninfo2, ninfo3)
    return (gr)

#print(combineGroups(thefirst, thesecond, thethird, thefourth).returngroupinfo())


##KEEP ALL OF THE BLOCK COMMENTED SECTIONS FOR NOW AND I MIGHT STILL USE THEM


# avsortedmembers=[]
# def addempty2(arr):
#     for x in range(4):
#         arr.append([])
# addempty2(avsortedmembers)
# #print (avsortedmembers)
# for i in members:
#     avsortedmembers[members[i].returnnumtimeslot()-1].append(members[i])
#
Groups = []
for x in range(10):
    totgroups = groupsfive[x] + groupsfour[x]
    Groups.append(totgroups)

#sort availabilities 4-9
#anastasias
#volunteers even first

def biggestGroup(a, y, z):
    Max = len(Groups[a])
    i = a
    if len(Groups[y]) > Max:
        Max =len(Groups[y])
        i=y
    if len(Groups[z]) > Max:
        Max = len(Groups[z])
        i=z
        if len(Groups[y]) > len(Groups[z]):
            Max = len(Groups[y])
            i=y
    return i ,Max


cantsort = []
SortedGroups = []
SortedMembers = []

while len(members) > 0:
    if members[0].returnnumtimeslot() == 0:
        cantsort.append(members.pop(0))
    else:
        avm = members[0].returnnumtimeslot()
        if len(Groups[avm-1]) > 0:
            index = avm-1
            size = 1
        elif avm ==1:
            index , size = biggestGroup(4,5,6)
        elif avm == 2:
            index, size = biggestGroup(4, 7, 8)
        elif avm ==3:
            index, size = biggestGroup(5, 7, 9)
        elif avm == 4:
            index, size = biggestGroup(6, 8, 9)

        if size == 0:
            cantsort.append(members.pop(0))
        else:
            SortedGroups.append(Groups[index].pop(0))
            SortedMembers.append(members.pop(0))
