import pandas as pd
import numpy as np
import math
import xlsxwriter
from string import Template
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

#--------   NEW
#anastasia     in terminal enter 'pip install easygui'    to run
import easygui
import tkinter as Tk
import xlwt
clientListAddress = ""
volunteerListAddress = ""
msg = "Please select files"
Tk.Tk().lift()
while True:
    if clientListAddress == "" and volunteerListAddress == "":
        msg = "Please select client or volunteer list"
    elif clientListAddress == "" or type(clientListAddress)=="class 'NoneType'":
        msg = "Please select client list"
    elif volunteerListAddress == "" :
        msg = "Please select volunteer list"
    else:
        msg = "Please click continue to match volunteers with clients or reselect files by clicking either list"
    choices = ["Client List", "Volunteer List", "Continue", "Cancel"]
    reply = easygui.buttonbox(msg, choices=choices)

    if reply == "Client List":
        clientListAddress = easygui.fileopenbox()
    elif reply == "Volunteer List":
        volunteerListAddress = easygui.fileopenbox()
    elif reply == "Cancel":
        break
    elif clientListAddress != "" and volunteerListAddress != "" and isinstance(volunteerListAddress, str)and isinstance(clientListAddress, str) and reply == "Continue":
        break
    elif reply == "Continue":
        msg = "File(s) not selected.  Please select a file before continuing"

# print (volunteerListAddress)
# print (clientListAddress)
# print (type(volunteerListAddress))
# print (reply)

# volunteer file

# xlsx = pd.ExcelFile('volunteers.xlsx')
xlsx = pd.ExcelFile(volunteerListAddress)
df = xlsx.parse(0)
df = df.drop(df.columns[[3, 6, 9, 12]], axis=1)

# members file

# members = pd.ExcelFile('members.xlsx')
members = pd.ExcelFile(clientListAddress)
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
    def returntime(self):
        return (self.time)
    def returnname1(self):
        return (self.name1)
    def returnname2(self):
        return (self.name2)
    def returnname3(self):
        return (self.name3)
    def returnname4(self):
        return (self.name4)
    def returnname5(self):
        return (self.name5)
    def returnemail1(self):
        return (self.email1)
    def returnemail2(self):
        return (self.email2)
    def returnemail3(self):
        return (self.email3)
    def returnemail4(self):
        return (self.email4)
    def returnemail5(self):
        return (self.email5)
    def returnav1(self):
        return (self.av1)
    def returnav2(self):
        return (self.av2)
    def returnif21(self):
        return (self.if21)
    def returnnetid(self):
        return (self.netid)
    def returninfo(self):
        return (self.info)

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
    def returnname(self):
        return self.name
    def returnphone(self):
        return self.phone
    def returnemail(self):
        return self.email
    def returnmethcontact(self):
        return self.methcontact
    def returnifcontacted(self):
        return self.ifcontacted
    def returnifconfirm(self):
        return self.ifconfirm
    def returntimeslot(self):
        return self.timeslot
    def returntask(self):
        return self.task
    def returnaddress(self):
        return self.address
    def returninfo(self):
        return self.info
    def returnflagged(self):
        return self.flagged



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

def checknan(teststring):
    if "nan" in teststring:
        return ""
    else:
        return teststring

def combineGroups(first, second, third = 0, fourth = 0):
    ## Determines order for greatest to least
    one = 0
    two = 0
    three = 0
    four = 0
    if first.returngroupsize() == 1 and second.returngroupsize() == 1 and third.returngroupsize() == 1 and fourth.returngroupsize() == 1:
        one = first
        two = second
        three = third
        four = fourth

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

    ## COMBINE 1 1 1 1
    if one.returngroupsize() == 1 and two.returngroupsize() == 1 and three.returngroupsize() == 1 and four.returngroupsize() == 1:
        nvol2name = two.returngroupinfo()[1]
        nvol2email = two.returngroupinfo()[2]
        nvol3name = three.returngroupinfo()[1]
        nvol3email = three.returngroupinfo()[2]
        nvol4name = four.returngroupinfo()[1]
        nvol4email = four.returngroupinfo()[2]

        nvol5name = float('nan')
        nvol5email = float('nan')

    ## COMBINE 3 and 2
    if one.returngroupsize() == 3 and two.returngroupsize() == 2:
        nvol2name = one.returngroupinfo()[3]
        nvol2email = one.returngroupinfo()[4]
        nvol3name = one.returngroupinfo()[5]
        nvol3email = one.returngroupinfo()[6]

        nvol4name = two.returngroupinfo()[1]
        nvol4email = two.returngroupinfo()[2]
        nvol5name = two.returngroupinfo()[3]
        nvol5email = two.returngroupinfo()[4]

    # COMBINE 2 and 2 OR 3 and 1
    if (one.returngroupsize() == 2 and two.returngroupsize() == 2) or (one.returngroupsize() == 3 and two.returngroupsize() == 1):
        nvol2name = one.returngroupinfo()[3]
        nvol2email = one.returngroupinfo()[4]
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

    # 2 1 1 or 2 1 1 1
    if (one.returngroupsize() == 2 and two.returngroupsize() == 1 and three.returngroupsize() == 1):
        nvol2name = one.returngroupinfo()[3]
        nvol2email = one.returngroupinfo()[4]
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

#----------------------------------------
# for y in range(10):
#     print (len(groupsthree[y]), len(groupstwo[y]), len(groupsone[y]))
#

# JUST 3 and 2
for y in range(10):
    if (len(groupsthree[y]) < len(groupstwo[y])):
        while len(groupsthree[y]) > 0:
            new = combineGroups(groupsthree[y][0], groupstwo[y][0])
            groupsfive.append(new)
            groupsthree[y].pop(0)
            groupstwo[y].pop(0)

    else:
       while len(groupstwo[y]) > 0:
           new = combineGroups(groupstwo[y][0], groupsthree[y][0])
           groupsfive.append(new)
           groupsthree[y].pop(0)
           groupstwo[y].pop(0)


## GROUPS 3s and 1s
for y in range(10):
    if (len(groupsthree[y])) < len(groupsone[y]):
        #for x in range(len(groupsthree[y])):
        while len(groupsthree[y]) > 0:
            new = combineGroups(groupsthree[y][0], groupsone[y][0])
            groupsfour.append(new)
            groupsthree[y].pop(0)
            groupsone[y].pop(0)

    else:
        while len(groupsone[y]) > 0:
            new = combineGroups(groupsone[y][0], groupsthree[y][0])
            groupsfour.append(new)
            groupsthree[y].pop(0)
            groupsone[y].pop(0)

## GROUPS 2s and 2s
for y in range(10):
    while (len(groupstwo[y])) > 1:
        new = combineGroups(groupstwo[y][0], groupstwo[y][1])
        groupsfour.append(new)
        groupstwo[y].pop(0)
        groupstwo[y].pop(0)

## GROUPS 2s and 1 and 1
for y in range(10):
    while len(groupstwo[y]) > 0 and len(groupsone[y]) >= 2:
        new = combineGroups(groupstwo[y][0], groupsone[y][0], groupsone[y][1])
        groupsfour.append(new)
        groupstwo[y].pop(0)
        groupsone[y].pop(0)
        groupsone[y].pop(0)

## GROUPS of 1 1 1 and 1
for y in range(10):
    while (len(groupsone[y])) >= 4:
        new = combineGroups(groupsone[y][0], groupsone[y][1], groupsone[y][2], groupsone[y][3])
        groupsfour.append(new)
        groupsone[y].pop(0)
        groupsone[y].pop(0)
        groupsone[y].pop(0)
        groupsone[y].pop(0)


# for y in range(10):
#     print (len(groupsthree[y]), len(groupstwo[y]), len(groupsone[y]))


### ALGORITHM PART STARTS HERE
Groups = []
Groups2 = []
for x in range(10):
    totgroups = groupsfive[x] + groupsfour[x]
    Groups2.append(totgroups)

#-----------------------------------------------------------------
Groups=[]

count=0
totgroups = []
for x in range(10):
    while len(groupsfive[x]) >0 or len(groupsfour[x]) >0:
        count=count +1
        if count%2 ==0 and len(groupsfive[x]) >0:
            totgroups.append(groupsfive[x].pop(0))
        elif len(groupsfour[x])>0:
            totgroups.append(groupsfour[x].pop(0))


    Groups.append(totgroups)
    totgroups = []


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
sortedDict = {}

while len(members) > 0:
    if members[0].returnnumtimeslot() == 0:
        cantsort.append(members.pop(0))
    else:
        avm = members[0].returnnumtimeslot()
        if len(Groups[avm-1]) > 0:
            index = avm-1
            size = 1
        elif avm ==1:
            index , size = biggestGroup(4, 5, 6)
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
            

<<<<<<< HEAD
# for x in sortedDict:
#     print ((sortedDict[x].returngroupinfo()), x.returnmeminfo())
   # print (sortedDict.keys((x.returngroupinfo())))
<<<<<<< HEAD

=======
>>>>>>> branch2
=======
>>>>>>> b0a62d09a5d5a71180c7e38c8e943f53e1c8f370

print (type(Groups[1][0].returntime()))

workbook = xlsxwriter.Workbook('Matched groups.xlsx')
worksheet = workbook.add_worksheet()
worksheet.write(0,0,"Community Member Name")
worksheet.write(0,1,"Phone Number")
worksheet.write(0,2,"Email")
worksheet.write(0,3,"Method of contact")
worksheet.write(0,4,"contacted to confirm?")
worksheet.write(0,5,"confirmed")
worksheet.write(0,6,"Date and time")
worksheet.write(0,7,"Task")
worksheet.write(0,8,"Adress")
worksheet.write(0,9,"Other Info")
worksheet.write(0,11,"Signup time")
worksheet.write(0,12,"Volunteer 1 Name")
worksheet.write(0,13,"Volunteer 1 Email")
worksheet.write(0,14,"Volunteer 2 Name")
worksheet.write(0,15,"Volunteer 2 Email")
worksheet.write(0,16,"Volunteer 3 Name")
worksheet.write(0,17,"Volunteer 3 Email")
worksheet.write(0,18,"Volunteer 4 Name")
worksheet.write(0,19,"Volunteer 4 Email")
worksheet.write(0,20,"Volunteer 5 Name")
worksheet.write(0,21,"Volunteer 5 Email")
worksheet.write(0,22,"Time 1")
worksheet.write(0,23,"Time 2")
worksheet.write(0,24,"Is someone over 21 and will drive?")
worksheet.write(0,25,"NetId")
worksheet.write(0,26,"Other info")
print(1)
row =1
for x in range(len(SortedGroups)):

    worksheet.write(row, 0, checknan(str(SortedMembers[x].returnname())))
    worksheet.write(row, 1, checknan(str(SortedMembers[x].returnphone())))
    worksheet.write(row, 2, checknan(str(SortedMembers[x].returnemail())))
    worksheet.write(row, 3, checknan(str(SortedMembers[x].returnmethcontact())))
    worksheet.write(row, 4, checknan(str(SortedMembers[x].returnifcontacted())))
    worksheet.write(row, 5, checknan(str(SortedMembers[x].returnifconfirm())))
    worksheet.write(row, 6, checknan(str(SortedMembers[x].returntimeslot())))
    worksheet.write(row, 7, checknan(str(SortedMembers[x].returntask())))
    worksheet.write(row, 8, checknan(str(SortedMembers[x].returnaddress())))
    worksheet.write(row, 9, checknan(str(SortedMembers[x].returninfo())))
    worksheet.write(row, 11, checknan(str(SortedGroups[x].returntime())))
    worksheet.write(row, 12,checknan(str(SortedGroups[x].returnname1())))
    worksheet.write(row, 13, checknan(str(SortedGroups[x].returnemail1())))
    worksheet.write(row, 14, checknan(str(SortedGroups[x].returnname2())))
    worksheet.write(row, 15, checknan(str(SortedGroups[x].returnemail2())))
    worksheet.write(row, 16, checknan(str(SortedGroups[x].returnname3())))
    worksheet.write(row, 17, checknan(str(SortedGroups[x].returnemail3())))
    worksheet.write(row, 18, checknan(str(SortedGroups[x].returnname4())))
    worksheet.write(row, 19, checknan(str(SortedGroups[x].returnemail4())))
    worksheet.write(row, 20, checknan(str(SortedGroups[x].returnname5())))
    worksheet.write(row, 21, checknan(str(SortedGroups[x].returnemail5())))
    worksheet.write(row, 22, checknan(str(SortedGroups[x].returnav1())))
    worksheet.write(row, 23,checknan(str( SortedGroups[x].returnav2())))
    worksheet.write(row, 24, checknan(str(SortedGroups[x].returnif21())))
    worksheet.write(row, 25, checknan(str(SortedGroups[x].returnnetid())))
    worksheet.write(row, 26, checknan(str(SortedGroups[x].returninfo())))
    row = row+1
row = row+2
print(1)
for x in range(len(cantsort)):
    worksheet.write(row, 0, checknan(str(cantsort[x].returnname())))
    worksheet.write(row, 1, checknan(str(cantsort[x].returnphone())))
    worksheet.write(row, 2, checknan(str(cantsort[x].returnemail())))
    worksheet.write(row, 3, checknan(str(cantsort[x].returnmethcontact())))
    worksheet.write(row, 4, checknan(str(cantsort[x].returnifcontacted())))
    worksheet.write(row, 5, checknan(str(cantsort[x].returnifconfirm())))
    worksheet.write(row, 6, checknan(str(cantsort[x].returntimeslot())))
    worksheet.write(row, 7, checknan(str(cantsort[x].returntask())))
    worksheet.write(row, 8, checknan(str(cantsort[x].returnaddress())))
    worksheet.write(row, 9, checknan(str(cantsort[x].returninfo())))
    row=row+1
print(1)
row = row+2
for y in range(len(Groups)):
    for x in range(len(Groups[y])):
        worksheet.write(row, 11, checknan(str(Groups[y][x].returntime())))
        worksheet.write(row, 12, checknan(str(Groups[y][x].returnname1())))
        worksheet.write(row, 13, checknan(str(Groups[y][x].returnemail1())))
        worksheet.write(row, 14, checknan(str(Groups[y][x].returnname2())))
        worksheet.write(row, 15, checknan(str(Groups[y][x].returnemail2())))
        worksheet.write(row, 16, checknan(str(Groups[y][x].returnname3())))
        worksheet.write(row, 17, checknan(str(Groups[y][x].returnemail3())))
        worksheet.write(row, 18, checknan(str(Groups[y][x].returnname4())))
        worksheet.write(row, 19, checknan(str(Groups[y][x].returnemail4())))
        worksheet.write(row, 20, checknan(str(Groups[y][x].returnname5())))
        worksheet.write(row, 21, checknan(str(Groups[y][x].returnemail5())))
        worksheet.write(row, 22, checknan(str(Groups[y][x].returnav1())))
        worksheet.write(row, 23, checknan(str(Groups[y][x].returnav2())))
        worksheet.write(row, 24, checknan(str(Groups[y][x].returnif21())))
        worksheet.write(row, 25, checknan(str(Groups[y][x].returnnetid())))
        worksheet.write(row, 26, checknan(str(Groups[y][x].returninfo())))
        row = row +1

workbook.close()

sortedLinks = []
sortedAddress = []

def createlink(inaddress):
    base = 'https://www.google.com/maps/dir/?api=1&origin=Beamish-Munro+Hall+ON&destination='
    a,b,c = inaddress.split(' ')
    target = a + '%20' + b + '%20' + c + '%20'+ 'Kingston%2C%20ON&travelmode=walking'
    outaddress = base + target
    return (outaddress)

for member in SortedMembers:
    addr = (member.returnmeminfo()[-3])
    sortedLinks.append(createlink(addr))
    sortedAddress.append(addr)

# WE HAVE SORTED LINKS
# CREATE EMAILS IN ORDER:
sortedEmails = []
sortedTime = []

for group in SortedGroups:
    email = (group.returngroupinfo()[2])
    sortedEmails.append(email)

for member in SortedMembers:
    otime = (member.returnmeminfo()[6])
    newtime = ""
    if otime == "Saturday Morning":
        newtime = "Saturday 9am - 12pm"
    elif otime == "Saturday Afternoon":
        newtime = "Saturday 1pm - 4pm"
    elif otime == "Sunday Morning":
        newtime = "Sunday 9am - 12pm"
    elif otime == "Sunday Afternoon":
        newtime = "Sunday 1pm - 4pm"

    sortedTime.append(newtime)

sortedEmails = ["anastasiavkrause@gmail.com", "16jmd9@queensu.ca","anastasiavkrause@gmail.com", "16jmd9@queensu.ca","anastasiavkrause@gmail.com", "16jmd9@queensu.ca","anastasiavkrause@gmail.com", "16jmd9@queensu.ca","anastasiavkrause@gmail.com", "16jmd9@queensu.ca","anastasiavkrause@gmail.com", "16jmd9@queensu.ca","anastasiavkrause@gmail.com", "16jmd9@queensu.ca","anastasiavkrause@gmail.com", "16jmd9@queensu.ca","anastasiavkrause@gmail.com", "16jmd9@queensu.ca","anastasiavkrause@gmail.com", "16jmd9@queensu.ca"]


def read_template(filename):
    with open(filename, 'r', encoding='utf-8') as template_file:
        template_file_content = template_file.read()
    return Template(template_file_content)

s = smtplib.SMTP(host= "smtp.gmail.com", port=587)
s.starttls()
s.login("akrause1999@gmail.com", "Anak1999!!")
#s = smtplib.SMTP_SSL('smtp.gmail.com:465')
#s.login('akrause1999@gmail.com', 'Anak1999!!')
#s.sendmail('from', 'to', msg.as_string())

message_template = read_template('message.txt')

for email, time, address, link in zip(sortedEmails, sortedTime, sortedAddress, sortedLinks):
    msg = MIMEMultipart()  # create a message

    # add in the actual person name to the message template
    message = message_template.substitute(TIME_SLOT=time.title(), ADDRESS= address.title(), LINK=link.title())

    # setup the parameters of the message
    msg['From'] = "akrause1999@gmail.com"
    msg['To'] = email
    msg['Subject'] = "Your Fix 'n' Clean Volunteer Assignment"

    # add in the message body
    msg.attach(MIMEText(message.lower(), 'plain'))

    # send the message via the server set up earlier.
    s.send_message(msg)
    del msg

s.quit()
