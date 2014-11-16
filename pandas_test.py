#!/usr/local/bin/python3.4
import pandas as pd
import numpy as np
import time

curMonth = int(time.strftime("%m"))
curYear = int(time.strftime("%Y"))

if curMonth > 6:
    curYear += 1

numberOfRows = ((curYear - 1900)*12)+6
pd.options.display.float_format = '{:,.2f}'.format

prng = pd.period_range('Jan-1900', periods=numberOfRows, freq='M')
indexFrame = pd.DataFrame(index=prng)


#df = pd.DataFrame(np.random.randint(100, size=len(prng)), prng, columns=['var2'])
df = pd.DataFrame(np.random.randint(100, size=numberOfRows), index=pd.date_range('01/01/1900',periods=numberOfRows,freq='MS'), columns=['var2'])
df = df.to_period()
print (df, df.index)

outputSeries = pd.concat([indexFrame,df], axis=1)

outputSeries['var2_m'] = 100*outputSeries['var2'].pct_change(periods=1, freq ='M')
outputSeries['var2_q'] = 100*outputSeries['var2'].pct_change(periods=12, freq ='M')

print (outputSeries)
