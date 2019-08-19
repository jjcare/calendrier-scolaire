# -*- coding: utf-8 -*-
#
# Python 3.3
#
#  MakeCNDEtudes.py  ---  2015, 2019 John Carey, CND
#
#     Parse files containing school days and study periods.
#     Produce ical file containing study periods for current year.
#
#     NOTE: input filename and schedule of études must be updated before use.
#
#
###################################################################################################################

import string, time, re, sys, datetime, random, codecs, os
import xlrd
import ical

def getSchoolYear():
    # return school year string
    thisyear = int(time.strftime('%Y')) - (int(time.strftime('%m')) < 5)  # consider before May as last year
    return '{:04d}-{:04d}'.format(thisyear, thisyear + 1)



class StudyIcal(ical.ICAL):  # local class based on ical class


    def __init__(self, thisyear):
        self.thisyear = thisyear
        try:
            self.jourlist = self.getJourlist()  # get the day numbers + dates as dict
            self.etudes = self.getEtudelist()   # get the schedule of study periods
            self.datelist = self.makeDatelist() # produce datelist for ical input
            ical.ICAL.__init__(self, self.datelist.split('\n'))   # prime ical class with data
        except:
            import dumper
            dumper.dump(self)  # do a big data dump!

    def getJourlist(self):
        fmname = 'jours école {}.tab'.format(self.thisyear)
        
        # get the school calendar and produce dictionary by day number (arbitrarily assigned)
        with open(fmname, encoding='utf8') as f:
            daylist = f.read().split('\n')  # read the data
            jourlist = {}
            # get the day numbers, reduced to unique values using a set
            jours = set([x.split('\t')[1] for x in daylist if x.strip() and x.split('\t')[1][0].isdigit()])
            # cycle through the days and capture all corresponding dates in a list
            for i in jours:
                dates = [ x.split('\t')[0] for x in daylist if x.strip() and x.split('\t')[1] == i]
                jourlist[i] = dates
        return jourlist

    def getEtudelist(self):
        # see if there is an excel sheet for the Grille_horaire

        fnames = [x for x in os.listdir() if re.match('Grille.*\.xlsx', x)]
        if fnames:

            # use excel method, finding most recent file
            fnames.sort(key=lambda x:os.stat(x).st_mtime, reverse=True)
            fname = fnames[0]
            print('{}| Using excel file: "{}"{}'.format(hr,fname,hr))
            
            # open sheet and process
            f = xlrd.open_workbook(fname)
            s = f.sheet_by_index(0)
            slots = [list(map(lambda x: x.value, s.row(i))) for i in range(s.nrows)]
            # get headings from heading line
            heads = [i for i in slots if 'Jour' in i[2]][0]
            # get locaux just in case later needed (not used currently)
            locaux = { ''.join(re.findall('([AB0-9])r?e? ',x[0])):x[0].split(' ')[-1]
                       for x in slots if re.match('\dr?e sec.*:',x[0])}
            # get the table of times and periods
            slots = [x for x in slots if re.match('\d{1,2} ?h ?\d{2}', x[0])]
            
        else:
            # use text file method - text must be joined into 8-col tabbed regions (for 6 day)
            inname = 'études {}.txt'.format(self.thisyear)
            print('{}| Using text file: "{}"{}'.format(hr,iname,hr))
            with open(inname, encoding='utf8') as f:
                buf = f.read().replace('"','').split('\n') # remove quotes and split lines
                heads = buf[0].split('\t')   # get column headings (days)
                jours = { j:[] for j in heads[1:-1] }  # get dict of days
                slots = [x.split('\t') for x in buf[1:] if x.strip()] # get periods
                
        # process timetable to get day dictionary
        jours = { j:[] for j in heads[1:-1] }  # get dict of days
        for i in slots:
            h1 = i[0] if i[0] else h1   # time, 1er cycle
            h2 = i[-1] if i[-1] else h2 # time, 2e cycle
            for j in range(1,len(heads)-1):
                if i[j].strip():
                    jours['Jour {}'.format(j)].append(self.getrv('/'.join(i[j].split('\n')), heads[j], h1, h2))

        return jours
        

    def getrv (self, act, jour, h1, h2):
        # return day, time, act
        hour = h1 if int(act[0]) < 3 else h2
        hour = re.sub(' ?h ?', ':', re.split(' ?[-–] ?',hour)[0])
        return '{}\t{}\t{}'.format(jour, hour, act.replace(' sec',''))

    def makeDatelist (self):
        # produce datelist for ical from self.jourlist and self.etudes
        ret = ''
        for i in self.jourlist:  # iterate through days
            for j in self.jourlist[i]: # get all dates for this schoolday
                for k in self.etudes['Jour {}'.format(i)]:  # iterate study periods
                    data = k.split('\t')
                    ret += '{}-{}\tÉtude {}\n'.format(j, data[1], data[2])
        return ret[:-1]



####### end of class StudyIcal(ical) #####

#---------------------------------------------
#            Main logic
#
#---------------------------------------------

# input and output files

datestring = getSchoolYear()

outname = 'CND Études {}.ics'.format(datestring)
hr = '\n|' + '-'*79

# begin processing

print ('{}| MakeCalEtude - produire ical version de l\'horaire des études.{}\n'.format(hr,hr))

random.seed('SHALOM')  # initialize random sequence for repeatability

etudes = StudyIcal(datestring) # get days and study periods in ical form

try:

    etudes.writefile(outname, echo=False)  # write the file

    print('{}| Terminé avec succès. {} événements ajoutés.{}'.format(hr,len(etudes.datelist.split('\n')),hr))
    print('| Résultat dans le fichier "{}".{}\n'.format(outname, hr))

except FileNotFoundError:
    print ('<<<<<<<<<<<<<<<<< E R R E U R >>>>>>>>>>>>>>>>>>>\n')
    print ( 'Une erreur est survenue. Assurez-vous que le fichier \n\n\t', inname)
    print ( '\nest disponible et relancez le programme.\n\n\n')
            
except:
    print ('<<<<<<<<<<<<<<<<<< E R R E U R >>>>>>>>>>>>>>>>>>\n')
    print ('Une erreur est survenue.')
    e = sys.exc_info()[0]
    print (e)

finally:
    print('Completed.')


        


