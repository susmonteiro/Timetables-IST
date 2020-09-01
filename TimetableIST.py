import xlwt, urllib.request, re
from xlwt import Workbook

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

#to be defined
MIN_HOUR = 8
MAX_HOUR = 19

subjects = {
    #subject : (url, colorTeoricas, colorPB. colorLabs)
    'BD' : ("https://fenix.tecnico.ulisboa.pt/disciplinas/BD225179577/2020-2021/1-semestre/turnos", 'sea_green', 'light_green', 'lime'),
    'CG' : ("https://fenix.tecnico.ulisboa.pt/disciplinas/CGra45179577/2020-2021/1-semestre/turnos", 'gold', 'ivory', 'light_yellow'),
    'IA' : ("https://fenix.tecnico.ulisboa.pt/disciplinas/IArt45179577/2020-2021/1-semestre/turnos", 'coral', 'rose', 'plum'),
    'OC' : ("https://fenix.tecnico.ulisboa.pt/disciplinas/OC11179577/2020-2021/1-semestre/turnos", 'aqua', 'light_blue', 'pale_blue'),
    'RC' : ("https://fenix.tecnico.ulisboa.pt/disciplinas/RC45179577/2020-2021/1-semestre/turnos", 'lavender', 'periwinkle', 'ice_blue')
}

def cleanup():
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

def getColor(subject, typeClass):
    if typeClass == 'T':
        return subjects[subject][1]
    elif typeClass == 'P':
        return subjects[subject][2]
    elif typeClass == 'L':
        return subjects[subject][3]
    else:
        return 'white'

def writeToExcel(classes, sheet):
    '''
    light_blue
    aqua
    '''
    weekIndex[0] = 0
    for k in range(1, 6):
        weekIndex[k] = weekIndex[k-1] + maxshift[k-1] + 1
    for c in classes:
        start, end, weekDay, subject, typeClass, shift = c.split('|')
        column = weekIndex[int(weekDay)] + int(shift) - 1
        style = xlwt.easyxf('align: horiz center, vert center;' 'pattern: pattern solid, fore_colour ' + getColor(subject, typeClass[0]) + '; border: left thin, right thin, top thin, bottom thin;')
        sheet.write_merge(int(start), int(end), column, column, subject + ' ' + typeClass, style)

def writeWeekDays(sheet):
    style = xlwt.easyxf('font: bold 1; align: horiz center; border: left medium, right medium, bottom thin')
    arr = ['Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta']
    i = 1
    for day in arr:
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
    patternGetClasses = " *<td>" + subject

    lines = info.splitlines()
    l = 0    #line

    for l in range(len(lines)):
        if re.search(patternGetClasses, lines[l]) and re.search(pattern, lines[l]):
            typeClass = re.findall(pattern, lines[l])[0]
            l += 2
            nameOfWeekDay = re.findall("S..|T..|Q..", lines[l])[0]
            dayOfWeek = getWeekDay(nameOfWeekDay)
            beginTime, finishTime = re.findall("\d{2}:\d{2}", lines[l])
            l += 5
            if re.search(' *LEIC-A', lines[l]):
                begin = timeCellNumber[beginTime]
                finish = timeCellNumber[finishTime]
                available = 1
                for i in range(begin, finish):
                    while str(i) + nameOfWeekDay + str(available) in globalclasses:
                        available += 1
                if available - 1 > maxshift[dayOfWeek]:
                    maxshift[dayOfWeek] = available - 1
                for i in range(begin, finish):
                    globalclasses.append(str(str(i) + nameOfWeekDay + str(available)))
                classes_code.append(str(begin) + '|' + str(finish-1) + '|' + str(dayOfWeek) + '|' + subject + '|' + typeClass + '|' + str(available))
        
def createSheet(sheet, pattern):
    createHours(sheet)
    for s in subjects:
        parseClasses(s, subjects[s][0], pattern)

    writeToExcel(classes_code, sheet)
    writeWeekDays(sheet)

###########         PROGRAMA            ##########

wb = Workbook()
sheetAll = wb.add_sheet('Tudo')
createSheet(sheetAll, "T\d{2}|L\d{2}|PB\d{2}")

cleanup()

sheetT = wb.add_sheet('Teoricas')
createSheet(sheetT, "T\d{2}")

wb.save('C:\\Users\smore\Desktop\horarioTeoricas.xls')


    





