# -*- coding: utf-8 -*-
"""
Created on Sun Mar  3 11:05:58 2019
@author: Z659190
"""

from os import chdir,listdir
from datetime import date
from pandas import DataFrame
from pandas import read_csv

import csv

#get path and filenames
root_dir = "C:/temp/s_data"
output_root = 'C:/temp/s_Data'
chdir(root_dir)

#get file list in the path
filenames=listdir(root_dir)
df = DataFrame()
df2 = DataFrame()
#open the target file, if not exist, create new one
now_date = date.today().strftime("%Y%m%d")
file2 = output_root + '/Original_result_'+now_date+'.csv'
for filename in filenames:
    if filename.startswith('New_Job'):
        filepath = root_dir+'/'+filename
        #for all files, read and process
        try:

            # df = DataFrame((read_csv(filepath_or_buffer=filepath,quoting=csv.QUOTE_NONE,delimiter=',',encoding='utf-8')))
            df = DataFrame((read_csv(filepath_or_buffer=filepath,delimiter=',',encoding='utf-8')))
            # df.loc[:,'filename'] = filename       #add new column
            # df.loc['new_line'] = '3'  "add new row
            df['filename'] = filename
            print(filepath)
            df2 = df2.append(df,sort=False)
        except Exception as e:
            print('Exception:', filepath,e)

df2.to_csv(file2,index=False)
print("File output to ",file2)
# sleep(10)




