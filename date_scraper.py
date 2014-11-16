#!/usr/local/bin/python3.4
from urllib.request import urlopen
import re
import time

#read in the javascript file from the ABS - this is updated monthly
for line in urlopen('http://abs.gov.au/AUSSTATS/abs@.nsf/webpages/ABS%20Release%20Calendar/$file/six_months_calendar.js'):
    lineStr = str( line, encoding='utf8' )
    if lineStr.startswith('AddEvent('):
        #convert the first space after the pub to a ~ seperator
        lineStr = lineStr.replace(" ","~",1)
        #Build the initial string replacements
        repls = ('AddEvent(', ''), (',\"','~'),('\"', ''),(');\r\n','')
        rep = dict((re.escape(k), v) for k, v in repls)
        pattern = re.compile("|".join(rep.keys()))
        #replace the patterns and split the line into a list
        data = pattern.sub(lambda m: rep[re.escape(m.group(0))], lineStr).split("~")
        #convert YYYYMMDD time into dd-mmm-YYYY
        pub_date = time.strftime("%d %B %Y",time.strptime(data[0], "%Y%m%d"))
        print (pub_date,data[1],data[2])
