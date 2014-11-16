 #!/usr/local/bin/python3.4
import time
import csv
import urllib.request, urllib.parse, urllib.error
import xlrd
import pandas as pd
import platform
#from xlutils.display import cell_display
#from xlrd.sheet import Cell

timerStart = time.time()
# build an empty Pandas dataframe to concatenate the retrieved data into
curMonth = int(time.strftime("%m"))
curYear = int(time.strftime("%Y"))

if curMonth > 6:
    curYear += 1

numberOfRows = ((curYear - 1900)*12)+6
pd.options.display.float_format = '{:,.2f}'.format

prng = pd.period_range('Jan-1900', periods=numberOfRows, freq='M')
indexFrame = pd.DataFrame(index = prng)

def fetchSeries(chapter, file, series):
    book = xlrd.open_workbook(file)
    # important - set xlrd datemode to match the excel book - avoiding any 1900/1904 problems
    datemode = book.datemode
    # build a temporary Pandas Series to hold the data
    s = pd.Series(name=chapter)
    if series[0]=="A":
        series = series + "_Data"

    # name_obj_list is a list of all named ranges in the excel book
    for nobj in book.name_obj_list:
        name = nobj.name
        if name == series:
            print (name)
            # result.value is the actual reference to the series (there can be multiple references - A1:A10,A18:A50 etc
            value = nobj.result.value
            for i in range(len(value)):
                # set Ref3d to the tuple of the form: (shtxlo, shtxhi, rowxlo, rowxhi, colxlo, colxhi)
                ref3d = value[i]
                for shx in range(ref3d.shtxlo, ref3d.shtxhi):
                    # go to the worksheet specified for the range
                    sh = book.sheet_by_index(shx)
                    # in case the range spans multiple columns
                    for colx in range(ref3d.colxlo, ref3d.colxhi):
                        # iterate over the rows
                        for rowx in range(ref3d.rowxlo, ref3d.rowxhi):
                            # value of the cell
                            cval = sh.cell_value(rowx, colx)
                            # get the date that corresponds to the series entry (column 0 is the date column)
                            cdate = sh.cell_value(rowx, 0)
                            # create a tuple holding the date
                            date_tup = xlrd.xldate_as_tuple(cdate, datemode)
                            # set the series value for this date
                            s[pd.datetime(date_tup[0], date_tup[1], date_tup[2])] = cval
    return s


links = set()

with open('MSB_Sources.csv', 'Ur') as f:
        datasources = list(tuple(rec) for rec in csv.reader(f, delimiter=','))
for item in datasources:
    links.add(tuple((item[1],item[2])))
for link in links:
    filename = urllib.request.urlretrieve(link[0], link[1])
    for item in datasources:
        if item[2] == link[1]:
            print (item)
            s = fetchSeries(item[0],item[2],item[3])
            s = s.to_period(freq='M')
            print (s)
            indexFrame = pd.concat([indexFrame, s], axis=1)

#indexFrame = indexFrame.groupby(indexFrame.index.month).sum()
print (indexFrame)

indexFrame.reindex_axis(sorted(indexFrame.columns), axis=1)
indexFrame.to_csv('./output.csv')
timerEnd = time.time()
timerInterval = timerEnd - timerStart
print("process took "+str(timerInterval)+" seconds")
print((platform.python_version()))
