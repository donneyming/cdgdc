# -*- coding: utf-8 -*-
"""
Created on Tue Nov  1 14:05:03 2016

@author: cms
"""
import pandas as pd
#from pandas import Series, DataFrame
import numpy as np
import re 

import time
import os
dirroot ='C:\\Users\\cms\\Desktop\\\学科评估\\问卷调查\\学校学科\\'

def compute():
    dfs = pd.DataFrame()

#文件合并
    for parent,dirnames,filenames in os.walk(dirroot): 
        for filename in filenames:
            file = dirroot + filename
            print (file)
            tmpdfs = pd.read_excel(file,converters={'PUBLISHNUMBER':int,'JOINNUMBER':int,'NOJOINNUMBER':int,'FINISHNUMBER':int,'XKZYDM':str})
            dfs = dfs.append(pd.DataFrame(tmpdfs))
#学校学科计算
#    print (dfs['XXDM'])
    dfs2 = dfs.groupby(['XXDM','XKZYDM','XXMC','XKZYMC']).sum()
    
    dfs_out = pd.DataFrame(dfs2)
    dfs_out['JOINPROPORTION'] = dfs_out['JOINNUMBER']/dfs_out['PUBLISHNUMBER'] 
    dfs_out['NOJOINPROPORTION'] = dfs_out['NOJOINNUMBER']/dfs_out['PUBLISHNUMBER']
    dfs_out['FINISHPROPORTION'] = dfs_out['FINISHNUMBER']/dfs_out['PUBLISHNUMBER']


    fileName = "学校学科"+re.sub(r'[^0-9]','',str(datetime.datetime.now()))  +".xls"  
    dfs_out.to_excel(fileName)

def main():
    compute()
if __name__ =='__main__':
    main()
