import time
import urllib
import xlrd
from xlutils.display import cell_display
from prettytable import PrettyTable
from xlrd.sheet import Cell

timerStart = time.time()
dataset_list = ('A163158W','A181479W')

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
    x = PrettyTable()
    x.add_column("Date", dataset_dict.keys())
    x.add_column(dataset, dataset_dict.values())
    x.sortby = "Date"
    print x
timerEnd = time.time()
timerInterval = timerEnd - timerStart
print "process took "+str(timerInterval)+" seconds"

#for key, value in sorted(dataset_dict.items()):
#    print("{} : {}".format(key, value))