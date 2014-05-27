import time
import urllib
import xlrd
import pandas as pd
#from xlutils.display import cell_display
#from xlrd.sheet import Cell

timerStart = time.time()

# Insert a comma seperated listof ABS Series IDs
dataset_list = ('A163158W','A181518A','A2325846C')

# build an empty Pandas dataframe to concatenate the retrieved data into
outputSeries = pd.DataFrame()

for dataset in dataset_list:
    # Using the methods outlined in http://www.abs.gov.au/AUSSTATS/abs@.nsf/97adb482c0aba769ca2570460017d0e7/15d1d9943893067dca2572ec00157f40!OpenDocument
    # append _Data to the dataset name to access only series data
    datasetRange = dataset+'_Data'
    # Call the most recent publication using the OpenAgent&sid= string and download it to tmp drive
    filename = urllib.urlretrieve("http://www.ausstats.abs.gov.au/ausstats/Meisubs.nsf/GetTimeSeries?OpenAgent&sid="+dataset, './'+dataset+'.xls')
    # use xlrd to open the workbook
    book = xlrd.open_workbook('./'+dataset+'.xls')
    # important - set xlrd datemode to match the excel book - avoiding any 1900/1904 problems
    datemode = book.datemode
    # build a temporary dictionary to hold the series data
    dataset_dict = {}

    # name_obj_list is a list of all named ranges in the excel book
    for nobj in book.name_obj_list:
        name = nobj.name
        if name == datasetRange:
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
                            # add the cell value to the dictionary, with the date tuple as the key
                            dataset_dict[xlrd.xldate_as_tuple(cdate, datemode)] = cval
    # build a temporary pandas dataframe with Date and dataset name columns
    s = pd.DataFrame(dataset_dict.items(), columns=['Date', dataset])
    # index the temporary dataframe on the date
    s = s.set_index('Date')
    # concatenate the temp dataframe to the master dataframe.  This will add a new column, creating new keys as required.
    outputSeries = pd.concat([outputSeries, s], axis=1)

outputSeries.to_csv('./output.csv')
print outputSeries
timerEnd = time.time()
timerInterval = timerEnd - timerStart
print "process took "+str(timerInterval)+" seconds"