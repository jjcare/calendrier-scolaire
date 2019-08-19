# -*- coding: utf-8 -*-
#
# Python 3.3
#
#  MakePCMRCal.py  ---  2015, John Carey, CND
#                       2017, rewritten for new school hours
#
#     Parse ical file containing school days and pcmr horaire.
#     Produce ical file containing pcmr practice periods.
#
#     NOTE: input filename pcmr.txt must be updated before use (copy-paste date lines from document)
#
#     Recurring events (e.g. tous les mardis) - added but Google calendar moves them at daylight savings
#
#
###################################################################################################################

import string, time, re, sys, datetime, random
import ical

def getSchoolYear():
    # return school year string
    now = time.localtime()
    thismonth = now[1]
    thisyear = now[0] - (thismonth < 5)  # consider before May as last year
    return '%4d-%4d' % (thisyear, thisyear + 1)


################################
# icalendar class (parse, print)
################################


class PCMR:

    daymatch = 'LUNDI|MARDI|MERCREDI|JEUDI|VENDREDI'
    MONTHS = 'dum,jan,fev,mar,avr,mai,jui,xxx,aou,sep,oct,nov,dec'.split(',')
    TIMEPAT = '(\d{1,2} ?[:h] ?\d{2})'   # matches 9:45, 9h45, 9 h 45, 09h45, 09 h 45, 09:45
    TIMESEQ = TIMEPAT + ' ?[/àÀ] ?' + TIMEPAT  # 14:05/15:05, 14:05 à 15:05, 15h55/16h55, etc
    NIVSEQ = '([1-5])r?e'
    DATEPAT = '(\d{1,2})[- ]([abcdefijlmnoprstuvéû]{3,5})'
    RECURRINGPAT = 'tous les '
    
    def __init__ (self, buf):
        # get the school year dates: YYYY-YYYY in a list
        self.yrs = list(map(int, re.search('\d{4}\-\d{4}',buf).group(0).split('-')))
        self.events, self.recurring = {}, {}
        self.processBuffer(buf.split('\n'))

    def processBuffer(self, buf):
        niveau, this_niveau, hour, group = [],[],'', ''
        for line in buf:
                line = line.strip()
                # check for hour
                have_hour = re.search(PCMR.TIMESEQ, line, re.I)
                if have_hour:
                    hour = self.getHour(line)[0] # for now, take start time and end will be +1hr

                # check for date
                have_date = re.search(PCMR.DATEPAT, line, re.I)
                # check for niveau
                have_niv = re.search(PCMR.NIVSEQ, line, re.I)

                # set the general niveau if not in the date line
                if not have_date and have_niv:  # no date, set both general and particular
                    niveau = this_niveau = self.getNiveau(line)
                    group = ''
                elif have_niv: # get it from date line
                    this_niveau = self.getNiveau(line)
                elif have_date:  # take it from general
                    this_niveau = niveau

                have_group = re.search('GROUPE \d',line, re.I)
                if have_group:
                    group = ' ({})'.format(have_group.group(0))

                # check for regular meetings
                have_recurring = re.search(PCMR.RECURRINGPAT, line, re.I)

                if have_hour and have_recurring:  # we have a recurring event
                    self.getRecurring (line, hour, niveau)
                    
                # if not recurring and we have a date, we can store the event info (checks if datetime already has events)
                elif have_date:
                    this_datetime = self.getDatetime (line, hour)
                    self.events[this_datetime] = self.events.get(this_datetime, []) + list(map(lambda x:x+group, this_niveau))
                    #print ('Niveau: {}, {}'.format(' et '.join(this_niveau),this_datetime))
                    

    def getMonth (self, mstring):
        # returns the month number of the month string given (e.g. 'sept' -> 9)
        mstring = mstring.translate(''.maketrans('éû','eu'))[:3]
        return PCMR.MONTHS.index(mstring)

    def getHour (self, hstring):
        # grab begin time and end time out of hstring (handles spaces and : or h/H
        h = re.search(PCMR.TIMESEQ, hstring, re.I)
        return list(map(lambda x:x.lower().replace('h',':').replace(' ',''), [h.group(1), h.group(2)]))

    def getDatetime (self, dstring, hour):
        # return a string 'YYYY-MM-DD-hh:mm'
        re_date = re.search(PCMR.DATEPAT, dstring, re.I)
        day, month = re_date.group(1), re_date.group(2)
        monthnum = self.getMonth(month)
        return '{}-{:02d}-{:02d}-{}'.format(self.yrs[monthnum < 8], monthnum, int(day), hour)

    def getNiveau (self, nstring):
        # grab niveau(x) from nstring - returns list
        return re.findall(PCMR.NIVSEQ,nstring,re.I)

    def getRecurring(self, line, hour, niveau):
        # store a recurring event (some tricky processing required
        jours = 'lundis mardis mercredis jeudis vendredis'.split()
        recur = re.search('tous les ([^ ]+)', line, re.I)
        # is it everyone or just e.g. Ténors et basses ?
        gsearch = re.search('tous les (.+?) /? ?(?:pour )?(?:les )?(.*)', line, re.I)
        group = gsearch.group(2).strip('/').strip().lower() if gsearch else 'tous'
        print('Recurring: niveau: {}, group:{}'.format(niveau, group or 'tous'))
        if recur:
            day = recur.group(1)
            jourindex = jours.index(day)
            date = self.getRecurringDate(line, hour, jourindex)
            self.recurring[date] = self.recurring.get(date, []) + ['sec. {} ({})'.format(' et '.join(niveau), group)]

    def getRecurringDate(self, line, hour, jourindex):
        # no date given. Use earliest so far as start
        wkdays = 'MO TU WE TH FR'.split()
        rday = wkdays[jourindex]
        sdate = [d for d in sorted(self.events)][0][:10]
        d = datetime.date(int(sdate[:4]), int(sdate[5:7]), int(sdate[8:10]))
        # shift date to proper weekday
        d = d + datetime.timedelta(days=(jourindex - d.weekday()))
        return '{}-{}R{}'.format(d.strftime('%Y-%m-%d'), hour, rday)

    def mkrecurring(self, day, info):
        # format recurring event line
        return '{}\tPCMR {}'.format(day, ', '.join(info))
    
    def mkevent (self, day, niveau):
        # format the event info for use in ical data structure
        return '{}\tPCMR {}'.format(day, ' et '.join([n[0] + ('re' if n[0] == '1' else 'e' if n[0].isdigit() else '') + n[1:] for n in niveau]))

    def getEvents (self):
        # convert dict of event dates to list for ical processing
        return [ self.mkevent(d,self.events.get(d)) for d in sorted(self.events) ] + [ self.mkrecurring(d, self.recurring.get(d)) for d in sorted(self.recurring)]


    
class Jours:
	jours = []
	def __init__ (self, fname):
		try:
			f = open(fname,'r', encoding='utf-8')
			self.jours = list(map(lambda x: x.split('\t'),re.split('\r?\n', f.read())))
		except:
			self.jours = []
			return False
	def getDay(self):
		for j in self.jours:
			yield(j)
	def getDate(self, d):
		for j in self.jours:
			if d == j[0]:
				return j
		return False

        

#---------------------------------------------
#            Main logic
#
#---------------------------------------------

# input and output files

datestring = getSchoolYear()

inname = 'pcmr.txt'
pcmrfname = 'PCMR %s.ics' % (datestring)
joursfname = 'jours école %s.tab' % (datestring)

# begin processing

print ('MakePCMRCal - produire ical version de l\'horaire des pcmr.\n\n')

random.seed('SASEC')  # initialize random sequence for repeatability

try:
    
    f = open(inname,'r', encoding='utf8')
    buf = f.read()
    f.close()

except FileNotFoundError:
    print ('<<<<<<<<<<<<<<<<< E R R E U R >>>>>>>>>>>>>>>>>>>\n')
    print ( 'Une erreur est survenue. Assurez-vous que le fichier \n\n\t', inname)
    print ( '\nest disponible et relancez le programme.\n\n\n')
            
except:
    print ('<<<<<<<<<<<<<<<<<< E R R E U R >>>>>>>>>>>>>>>>>>\n')
    print ('Une erreur est survenue.')
    e = sys.exc_info()[0]
    print (e)

# instantiate PCMR class with input buffer
pcmr = PCMR(buf)

cal = ical.ICAL(pcmr.getEvents())

f = open(pcmrfname, 'w', encoding='utf-8')
f.write(cal.ical)
f.close()
print('Processing completed. Calendar in {}.'.format(pcmrfname))
print('Remember to correct recurring events by setting the timezone.')

