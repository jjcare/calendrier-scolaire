# assignschooldays.py
#
# make school calendar ical entries (calculated - see below)
#
#
#  John Carey, 2017, 2019
#     2019: gui version using tkinter (relies on Python 3.7
#

from tkinter import *
import time, datetime, os.path, re
from datetime import date
import ical   # module for preparing icalendar format

docstring = """
[--------------------- assignschooldays.py ---------------------]

This program creates the school calendar automatically based on
the current year, Quebec holidays and input of special days.

March break and Easter week (including Good Friday) are included
automagically.

Special days must be included in a config file called specialsYYYY.txt
(where YYYY is the current schoolyear). Day types are given in [], followed
by one date per line.

Dates are entered in YYYY-MM-DD format. After the first whole date
only the month or even the day can be given if the same as previous
date.

Excluding weekends, holidays and special days, the program then
assigns days 1 to 6 to the remaining dates.

[----------------------------------------------------------------]

"""
fields = { "August":"First School day in August",
           "June":"Last School day in June",
           "January":"First school day in January"}

class SchoolYear:
    
    def __init__(self):
        self.thisyr = date.today().year # - (date.today().month < 5)  # back a year after New Year

    def processYear(self, indays=None):
        self.selects = indays
        self.dates = self.getDates()
        self.marchbreak()
        self.ferias = self.getQuebecHolidays()
        self.markEasterWeek()
        self.markHolidays()
        if os.path.exists('specials{}.txt'.format(self.thisyr)):
            self.getSpecialsFromFile()
        else:  # interactive entry
            for t in ['rentrée','journée pédagogique','examens']:
                self.markOffDays(t)
        print('\nSetting days of school week...', end=' ')
        self.addSchoolDays()
        print('Done.')

        
    def getDates(self):
        s = self.selects["August"].get()
        e = self.selects["June"].get()
        try:
            start, end = (date(self.thisyr,8, int(s)),
                          date(self.thisyr+1, 6, int(e)))
            print('Start: {}, End: {}\n'.format(start,end))
        except:
            print('Invalid input data. Cancelling program.', file=sys.stderr)
            raise IndexError

        days = [ start + datetime.timedelta(days=x) for x in range((end-start).days + 1)]

        # convert to desired format, eliminating weekends
        days = [ i.strftime('%Y-%m-%d\t%w\t') for i in days if i.strftime('%w') not in '06']

        # get Christmas holidays
        ret = self.selects["January"].get()
        for i, day in enumerate(days):
            if day[5:10] == '12-25': break
        while days[i][5:10] != '01-{:02d}'.format(int(ret)):
            days[i] += 'Congé' if days[i][-1] == '\t' else ''
            i += 1
        return days

    def marchbreak (self):
        # set the March break
        firstmarch = self.dates.index([i for i in self.dates if '-03-' in i][0]) # get first day
        while self.dates[firstmarch].strip()[-1] != '1':  # find Monday
            firstmarch += 1
        for i in range(firstmarch, firstmarch+5):   # set five days as congé
            self.dates[i] += 'Congé'
        print('March break calculated: {} - {}'.
              format(self.dates[firstmarch][:10],
                     self.dates[firstmarch+4][:10]))

    def markEasterWeek (self):
        # get Easter Monday - must be called after getQuebecHolidays
        easterMonday = [i for i in sorted(self.ferias) if 'Easter Monday' in self.ferias[i]][1]
        # find position in date list
        firstd = self.dates.index([i for i in self.dates if easterMonday in i][0])
        if easterMonday > str(self.thisyr + 1) + '-04-19':
            firstd -= 4 # back up a week if Easter is late
            print ("Late Easter - assigning Holy Week as congé.")
            
        # now set whole week as congé
        for i in range(firstd-1, firstd+5): # include Good Friday
            self.dates[i] += 'Congé' if self.dates[i][-1] == '\t' else ''  # don't overwrite

    def getQuebecHolidays (self):
        # get Quebec holidays and mark as congés
        import holidays
        ferias = { str(da):name for da,name in
                   sorted(holidays.CA(prov='QC', years=[self.thisyr,self.thisyr+1]).items()) }
        return ferias
        
    def markHolidays (self):
        # set congé for holidays
        for i,day in enumerate(self.dates):
            if day[:10] in self.ferias:
                self.dates[i] += 'Congé ferié' if self.dates[i][-1] == '\t' else ''

    def answers (self, daytype):
        print('\nInput dates for type "{}" (blank line ends).'.format(daytype))
        ans = input('?')
        while ans:
            yield ans
            ans = input('?')

    def markOffDays (self, daytype):
        # use for ped days and exam days ('journée pédagogique', '0')
        answer, ped = '**', []
        save = ['','','']
        print('\nInput "{}" for this year (blank line ends).'.format(daytype))
        answer = input('?')
        while answer:  # continue until blank line
            indate = answer.split('-')
            save[-(len(indate)):] = indate  # update least significant date parts
            ped.append('-'.join(save))  # no error checking !!
            answer = input('?')
            
        for i,day in enumerate(self.dates):
            if day[:10] in ped:
                self.dates[i] += daytype if self.dates[i][-1] == '\t' else ''

    def getSpecialsFromFile(self):
        # use input file for ped days etc.
        # format is [daytype], followed by date lines (eg. 2018-01-01, 02-04, 26)
        data = open('specials{}.txt'.format(self.thisyr),encoding='utf-8-sig').read()
        items = re.split(r'\[(.+?)\]', data)[1:]  # split sections and chop off null value
        specials = dict(zip(*[iter(items)]*2)) # - no copies made
        # now process the sections
        for daytype in specials:
            dlist = specials[daytype].strip().split('\n')
            ped, save = [], ['','','']
            for day in dlist:
                indate = day.split('-')
                save[-(len(indate)):] = indate  # update least significant part
                ped.append('-'.join(save)) # put date back together
            # apply dates to calendar
            for i,d in enumerate(self.dates):
                if d[:10] in ped:
                    self.dates[i] += daytype if self.dates[i][-1] == '\t' else ''
        print('Special dates updated from file specials{}.txt.'.format(self.thisyr))
        

    def addSchoolDays(self):
        # consume input lines and add days (1 to 6, but could be changed)
        jour = 0
        for i,d in enumerate(self.dates):
            if d[-1] == '\t':  # check for special already assigned to date
                self.dates[i] += 'JOUR ' + str(jour % 6 + 1)
                jour += 1

    def showDate(self, day):
        print( self.dates[self.dates.index([i for i in self.dates if day in i][0])])

    def __str__(self):
        """str rep -> YYYY-MM-DD\tjour"""
        ret = ''
        for d in self.dates:
            inf = d.split('\t')
            jour = inf[2][-1] if re.match('JOUR [1-6]', inf[2]) else inf[2]
            ret += '{}\t{}\n'.format(inf[0], jour)
        return ret
    

def makeform(root, fields):
   entries = {}
   for field in fields:
      row = Frame(root)
      lab = Label(row, width=25, text=fields[field]+": ", anchor='e')
      ent = Entry(row)
      ent.insert(0,"")
      row.pack(side = TOP, fill = X, padx = 5 , pady = 5)
      lab.pack(side = LEFT)
      ent.pack(side = LEFT, expand = YES)
      entries[field] = ent
   return entries


def make_schoolyear(entries):
    # received values, make a school year file
    yr = SchoolYear()
    yr.processYear(entries)

    # now make ical from school calendar
    
    cal = ical.ICAL(yr.dates)
    outfile = 'Calendrier Scolaire {}-{}.ics'.format(yr.thisyr, yr.thisyr+1)
    f = open(outfile, 'w', encoding='utf-8')
    f.write(cal.getCalendar())
    f.close()
    tabfile = 'jours école {}-{}.tab'.format(yr.thisyr, yr.thisyr+1)
    f = open(tabfile, 'w', encoding='utf-8')
    f.write(str(yr))
    f.close()
    print('Calendar has been written to {}, {}.'.format(outfile, tabfile))
    label = Label(root, text='Calendar has been written to {}, {}.'.format(outfile, tabfile))
    label.pack()


###  M A I N  L O G I C

if __name__ == '__main__':
    root = Tk()
    root.title("Make School Calendar")
    msg = Message(root, text=docstring)
    msg.pack(padx = 5)
    ents = makeform(root, fields)  # make the labels
    #root.bind('<Return>', (lambda event, e = ents: fetch(e)))
    b1 = Button(root, text = 'Calculate School Year',
      command=(lambda e = ents: make_schoolyear(e)))
    b1.pack(side = LEFT, padx = 5, pady = 10)
    b3 = Button(root, text = 'Quit', command = root.quit)
    b3.pack(side = LEFT, padx = 5, pady = 5)
    root.mainloop()

        
