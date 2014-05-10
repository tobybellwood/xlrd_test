import urllib
import xlrd
filename = urllib.urlretrieve("http://www.ausstats.abs.gov.au/ausstats/ABS@Archive.nsf/GetTimeSeries?OpenAgent&sid=A2519058W")
book = xlrd.open_workbook(filename)                                                                                                                                                                               
for nobj in book.name_obj_list                                                                                                                                                                                            
	name = nobj.name
	res = nobj.result
  kind = res.kind
  value = res.value
  print name,res,kind,value
