import pandas as pd
import xlrd
import datetime

d = pd.ExcelFile('C:\Users\helen\Desktop\January.xlsx')
for s in d.sheet_names:
    df = d.parse(s,skiprows = 1)
    for i in range(df.shape[0]):
        if df.SHIPPER[i] == 'NTUC':
            if df.LOCATION[i] == 'V':
                if str(df.REMARKS[i]) != 'nan':
                    print '\033[1;32m %s' %s
                    print '\033[1;32m %s' %df['CONTR NUMBER'][i]
                    print '\033[1;32m %s' %df.LOCATION[i]
                    print '\033[1;32m %s' %df.REMARKS[i]
                    print '\n'
            elif df.LOCATION[i] == 'F' or df.LOCATION[i] == 'D':
                if str(df.REMARKS[i]) != 'nan':
                    print '\033[1;35m %s' %s
                    print '\033[1;35m %s' %df['CONTR NUMBER'][i]
                    print '\033[1;35m %s' %df.LOCATION[i]
                    print '\033[1;35m %s' %df.REMARKS[i]
                    print '\n'
                    
 for s in d.sheet_names:
    df = d.parse(s,skiprows = 1)
    for i in range(df.shape[0]):
        if df.SHIPPER[i] == 'XXXX':
            if df.LOCATION[i] == 'F':
                 if not isinstance (df.REMARKS[i], basestring):
                    start = datetime.datetime.strptime(df.DATE[i],'%d/%m')
                    end = datetime.datetime.strptime(df['DATE.1'][i], '%d/%m')
                    if end-start > pd.to_timedelta(4, unit = 'd'):
                        print '\033[1;30m %s' %s
                        print '\033[1;30m %s'%df['CONTR NUMBER'][i]
                        print '\033[1;30m %s' %(start.strftime('%d/%m'))
                        print '\033[1;30m %s' %(end.strftime('%d/%m'))
                        print '\033[1;30m %s' %(end - start + pd.to_timedelta(1, unit = 'd'))
                        print '\033[1;30m %s' %df.LOCATION[i]
                        print '\n'
                    else:
                        print '\033[1;30m %s' %s
                        print '\033[1;30m %s'%df.LOCATION[i]
                        print '\033[1;30m %s' %df['CONTR NUMBER'][i]
                        print '\033[1;30m %s' %df.REMARKS[i]
                        print '\033[1;30m %s' %df['DATE.1'][i]
                        print '\n'
