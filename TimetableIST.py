import xlwt, urllib.request, re
Workbook = xlwt.Workbook

# global things
timeCellNumber = dict()             #each half hour has a cell in excel  
maxshift = {                    #how many columns each weekday needs
    0 : 0,
    1 : 0,
    2 : 0,
    3 : 0,
    4 : 0,
    5 : 0
}
weekIndex = dict()              #first column of each weekday
globalclasses = []              #code for each class to check conflit of columns
classes_code = []               #ready to be written to the excel file

MIN_HOUR = 8
MAX_HOUR = 19

subjects = {
    'BD' : {
        'url' : "https://fenix.tecnico.ulisboa.pt/disciplinas/BD225179577/2020-2021/1-semestre/turnos", 
        'T' : 'sea_green', 
        'P' : 'light_green',
        'L' : 'lime'
        },
    'CG' : {
        'url' : "https://fenix.tecnico.ulisboa.pt/disciplinas/CGra45179577/2020-2021/1-semestre/turnos",
        'T' : 'gold',
        'P' : 'ivory',
        'L' : 'light_yellow'
        },
    'IA' : {
        'url' : "https://fenix.tecnico.ulisboa.pt/disciplinas/IArt45179577/2020-2021/1-semestre/turnos",
        'T' : 'coral',
        'P' : 'rose',
        'L' : 'plum'
        },
    'OC' : {
        'url' : "https://fenix.tecnico.ulisboa.pt/disciplinas/OC11179577/2020-2021/1-semestre/turnos",
        'T' : 'aqua',
        'P' : 'light_blue',
        'L' : 'pale_blue'
        },
    'RC' : {
        'url' : "https://fenix.tecnico.ulisboa.pt/disciplinas/RC45179577/2020-2021/1-semestre/turnos",
        'T' : 'lavender',
        'P' : 'periwinkle',
        'L' : 'ice_blue'
        }
}

def init():
    timeCellNumber.clear()
    for k in maxshift:
        maxshift[k] = 0
    weekIndex.clear()
    globalclasses.clear()
    classes_code.clear()

def getWeekDay(day):
    weekDay = {
        'Seg': 1,
        'Ter': 2,
        'Qua': 3,
        'Qui': 4,
        'Sex': 5
    }
    return weekDay.get(day)

def writeToExcel(classes, sheet):
    weekIndex[0] = 0
    for k in range(1, 6):
        weekIndex[k] = weekIndex[k-1] + maxshift[k-1] + 1
    for c in classes:
        column = weekIndex[c['weekDay']] + c['shift'] - 1
        style = xlwt.easyxf('align: horiz center, vert center;' 'pattern: pattern solid, fore_colour ' + subjects[c['subject']][c['typeClass'][0]] + '; border: left thin, right thin, top thin, bottom thin;')
        sheet.write_merge(c['start'], c['end'], column, column, c['subject'] + ' ' + c['typeClass'], style)

def writeWeekDays(sheet):
    style = xlwt.easyxf('font: bold 1; align: horiz center; border: left medium, right medium, bottom thin')
    i = 1
    for day in ['Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta']:
        sheet.write_merge(0, 0, weekIndex[i], weekIndex[i] + maxshift[i], day, style)
        i+=1

def createHours(sheet):
    hour = MIN_HOUR
    minute = '00'
    i = 1

    while hour < MAX_HOUR:
        hourstr = '0' + str(hour) if hour < 10 else str(hour)
        sheet.write(i, 0, hourstr + ':' + minute)
        timeCellNumber[hourstr + ':' + minute] = i
        if minute == '00':
            minute = '30'
        else:
            minute = '00'
            hour+=1
        i+=1

def getPageCode(url):
    # get pagecode from url
    fp = urllib.request.urlopen(url)
    page_source = fp.read()
    whole = page_source.decode("utf8")
    fp.close()
    return whole

def parseClasses(subject, url, pattern):
    info = getPageCode(url)
    patternGetClasses = " *<td>(a_)?" + subject

    lines = info.splitlines()

    for i,l in enumerate(lines):
        if re.search(patternGetClasses, l) and re.search(pattern, l):
            typeClass = re.findall(pattern, l)[0]
            i += 2

            nameOfWeekDay = re.findall("S..|T..|Q..", lines[i])[0]
            dayOfWeek = getWeekDay(nameOfWeekDay)

            beginTime, finishTime = re.findall("\d{2}:\d{2}", lines[i])
            i += 5

            # if re.search(' *LEIC-A', lines[i]):   #switch with following line for leic
            while (not re.search("</td>", lines[i])):   #when the class belongs to multiple "teams"
                if re.search(' *MEIC-A', lines[i]):
                    begin = timeCellNumber[beginTime]
                    finish = timeCellNumber[finishTime]
                    available = 1
                    for idx in range(begin, finish):
                        while str(idx) + nameOfWeekDay + str(available) in globalclasses:
                            available += 1
                    if available - 1 > maxshift[dayOfWeek]:
                        maxshift[dayOfWeek] = available - 1
                    for idx in range(begin, finish):
                        globalclasses.append(str(str(idx) + nameOfWeekDay + str(available)))
                    classes_code.append({'start' : begin, 'end' : finish-1, 'weekDay' : dayOfWeek, 'subject' : subject, 'typeClass' : typeClass, 'shift' : available})
                i+=1

                 
def createSheet(sheet, pattern):
    init()
    createHours(sheet)
    for s in subjects:
        parseClasses(s, subjects[s]['url'], pattern)

    writeToExcel(classes_code, sheet)
    writeWeekDays(sheet)

###########         PROGRAMA            ##########
print("Starting...")

wb = Workbook()
sheetAll = wb.add_sheet('Tudo')
createSheet(sheetAll, "T\d{2}|L\d{2}|PB\d{2}")

sheetT = wb.add_sheet('Teoricas')
createSheet(sheetT, "T\d{2}")

wb.save('/mnt/c/Users/Susana/Dropbox/IST/masters/horario.xls')

print("Done!")