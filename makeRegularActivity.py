# -*- coding: utf-8 -*-
#
#  makeRegularActivity - produces ical entries for regular meetings (days 1-6)
#   John Carey, September 2011, version 1.0
#               August 2015, version 2.0 (period given manually - variable études)
#               September 2017, version 2.5 (imports ical routines, takes begin and end dates)
#
#   Procedure: Ask for activity name, level and day (as tuple?).
#              Output calendar file [Activity]-[Year].ics
#              Dependency: "jours école YYYY-YYYY.tab" - list of dates and jours (school calendar)
#
################# General defs  ################

import string, time, re, sys, datetime, random
from datetime import date
import ical

# get the current school year - i.e. this year, back a year after New Year until May
thisyr = date.today().year - (date.today().month < 5)  

# get this years school calendar filename
schoolcalfile = 'jours école {}-{}.tab'.format(thisyr, thisyr+1)

# data structure for activity information
class Info:
    """Information on regular activity"""

    def __init__ (self, s):
        self.items = re.split(', ?', s)
        if len(self.items) < 4: return None  # some info is missing
        self.act = self.items[0]
        self.jour = self.items[1]
        self.heure = self.items[2].replace('h',':')
        self.start = self.mkdate(self.items[3])
        # get end date or one year from begin date
        self.end = self.mkdate(self.items[4]) if len(self.items) > 4 else self.start.replace(year=self.start.year+1)

    def mkdate (self, day):
        # return a datetime object for this day (used for calculations)
        if isinstance(day, str) and day != "":
            y,m,d = day.split('-')
            return datetime.date(int(y), int(m), int(d))
        return None

    def makeEvent (self, date):
        # format the event string for ICAL use
        return '{}-{}\t{}'.format(date, self.heure, self.act)
       
    def inDateRange (self, day):
        # verify that the date is within the school calendar year
        check = self.mkdate(day)
        return self.start <= check <= self.end

        
        
# MAIN LOGIC - Ask for activity info: Name, starting date, day num, level
#

#random.seed('events')  # initialize random sequence for repeatability

# prompt message for activities
instring = "Enter activity name, day, heure, starting date, ending date(y-m-d): "

# Get the school calendar                
f = open(schoolcalfile, 'r', encoding='utf-8')
schooldays = f.read().split('\n')
f.close()

# print intro greeting
print ("          Make Regular Activity Calendar\n")
print ("Produce calendar events based on activity information.\n")

# loop while input is given
while True:
    inp = input (instring)
    if inp == '': break
    info = Info(inp)
    if not info:
        print ("Please separate items with commas.")
        continue

    # create list of dates that correspond to the data entered
    datelist = []
    for i in schooldays:
        if i.strip():
            li = i.split('\t')
            # check for matching date: in range and having jour number
            if info.inDateRange(li[0]) and re.match(info.jour, li[1]):
                datelist.append(info.makeEvent(li[0]))
                print (li)
    if not datelist:
        print ("Sorry, no events created.\n")
        continue

    # create ical version of datelist and write to output file
    cal = ical.ICAL(datelist)
    fname = '{}-{}.ics'.format(re.sub(r'[/\;:+&?!.]', '-', info.act), thisyr)
    cal.writefile(fname)
    
    print ("{} events created. Calendar events in file {}.\n".format(len(datelist), fname))

print ("\nProgram completed. Thanks for using Shalom informatique.\n")


    

