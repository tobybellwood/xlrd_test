import time
import urllib
import xlrd
import pandas as pd
from xlutils.display import cell_display
from xlrd.sheet import Cell

timerStart = time.time()
dataset_list = ('A163158W','A181518A','A2325846C')
outputSeries = pd.DataFrame()

for dataset in dataset_list:
    datasetRange = dataset+'_Data'
    filename = urllib.urlretrieve("http://www.ausstats.abs.gov.au/ausstats/Meisubs.nsf/GetTimeSeries?OpenAgent&sid="+dataset, '/tmp/'+dataset+'.xls')
    book = xlrd.open_workbook('/tmp/'+dataset+'.xls')
    datemode = book.datemode
    dataset_dict = {}

    for nobj in book.name_obj_list:
        name = nobj.name
        if name == datasetRange: #or name == 'Date_Range':
            value = nobj.result.value
            for i in range(len(value)):
                ref3d = value[i]
                for shx in range(ref3d.shtxlo, ref3d.shtxhi):
                    sh = book.sheet_by_index(shx)
                    for rowx in range(ref3d.rowxlo, ref3d.rowxhi):
                        for colx in range(ref3d.colxlo, ref3d.colxhi):
                            cty = sh.cell_type(rowx, colx)
                            cval = sh.cell_value(rowx, colx)
                            cdate = sh.cell_value(rowx, 0)
                            if cty == 3:
                                cval = cell_display(Cell(cty, cval))
                                #sval = xlrd.xldate_as_tuple(cval,datemode)
                            dataset_dict[xlrd.xldate_as_tuple(cdate, datemode)] = cval
    s = pd.DataFrame(dataset_dict.items(), columns=['Date', dataset])

    s = s.set_index('Date')
    outputSeries = pd.concat([outputSeries, s], axis=1)

outputSeries.to_csv('/tmp/output.csv')
print outputSeries
timerEnd = time.time()
timerInterval = timerEnd - timerStart
print "process took "+str(timerInterval)+" seconds"