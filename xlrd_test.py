import urllib
import xlrd
from xlrd.timemachine import REPR

def showable_cell_value(celltype, cellvalue, datemode):
    if celltype == xlrd.XL_CELL_DATE:
        try:
            showval = xlrd.xldate_as_tuple(cellvalue, datemode)
        except xlrd.XLDateError as e:
            showval = "%s:%s" % (type(e).__name__, e)
    elif celltype == xlrd.XL_CELL_ERROR:
        showval = xlrd.error_text_from_code.get(
            cellvalue, '<Unknown error code 0x%02x>' % cellvalue)
    else:
        showval = cellvalue
    return showval

dataset = 'A2519058W'
filename = urllib.urlretrieve("http://www.ausstats.abs.gov.au/ausstats/ABS@Archive.nsf/GetTimeSeries?OpenAgent&sid="+dataset, '/tmp/'+dataset+'.xls')
book = xlrd.open_workbook('/tmp/'+dataset+'.xls')
for nobj in book.name_obj_list:
    name = nobj.name
    if name == dataset:
        res = nobj.result
        kind = res.kind
        value = res.value
        print name,res,kind,value
        for i in range(len(value)):
            ref3d = value[i]
            print("Range %d: %s ==> %s"% (i, REPR(ref3d.coords), REPR(xlrd.rangename3d(book, ref3d))))
            datemode = book.datemode
            for shx in range(ref3d.shtxlo, ref3d.shtxhi):
                sh = book.sheet_by_index(shx)
                print("   Sheet #%d (%s)" % (shx, sh.name))
                rowlim = min(ref3d.rowxhi, sh.nrows)
                collim = min(ref3d.colxhi, sh.ncols)
                for rowx in range(ref3d.rowxlo, rowlim):
                    for colx in range(ref3d.colxlo, collim):
                        cty = sh.cell_type(rowx, colx)
                        cval = sh.cell_value(rowx, colx)
                        sval = showable_cell_value(cty, cval, datemode)
                        print("      (%3d,%3d) %-5s: %s"
                            % (rowx, colx, xlrd.cellname(rowx, colx), REPR(sval)))
