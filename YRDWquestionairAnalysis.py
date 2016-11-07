# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""

import pandas as pd
from pandas import Series,DataFrame

import numpy as np
import re 
import time
import os
import os.path
import datetime


dirroot = u"C:\\Users\\cms\\Desktop\\学科评估\\问卷调查\\回收\\问卷【研究生职业发展调查问卷】调查结果明细表\\"
excelroot = u"C:\\Users\\cms\\Desktop\\学科评估\\问卷调查\\模板\\下拉列表导入.xlsx"




NAMEARRAY = [["农、林、牧、渔业","采矿业","制造业","电力、热力、燃气及水生产和供应业","建筑业","批发和零售业","交通运输、仓储和邮政业","住宿和餐饮业","信息传输、软件和信息技术服务业","金融业","房地产业","租赁和商务服务业","科学研究和技术服务业","水利、环境和公共设施管理业","居民服务、修理和其他服务业","教育","卫生和社会工作","文化、体育和娱乐业","公共管理、社会保障和社会组织","国际组织","军队","其他"],
            ["业务工作“直接上级”","部门领导","人力资源部门人员","本部门同事","其他部门同事","其他"],
            ["很不熟悉","不熟悉","一般","熟悉","很熟悉"],
            ["1","2","3","4","5","6"],
            ["1","2","3","4","5","6"],
            ["1","2","3","4","5","6"],
            ["1","2","3","4","5","6"],
            ["专业基础知识","职业道德素养","心理调适能力","自我学习能力","实践动手能力","创新思维能力","沟通交流能力","组织管理能力","团队协作能力","工作态度","其他"],
            ["很不满意","不满意","不太满意","基本满意&nbsp;","满意","很满意"],
            ["很不愿意","不愿意","一般","愿意","很愿意"],
            [""],
            [""],
            [""]
            ]            
            
QUEARRAY = ["贵单位所属的行业是（下拉列表）",
            "您是该毕业生所在单位的",
            "您对该毕业生的熟悉程度是",
            "⑴岗位所需的专业知识",
            "⑵职业道德（含工作态度等）",
            "⑶工作能力",
            "⑷工作业绩",
            "贵单位招聘毕业生时，最注重哪些能力素养",
            "您对该毕业生的总体评价是",
            "贵单位是否愿意继续招聘该校毕业生?",
            "请列出您最愿意招聘毕业生的三所院校（含科研院所）名称",
            "请列出您最愿意招聘毕业生的三所院校（含科研院所）名称",
            "请列出您最愿意招聘毕业生的三所院校（含科研院所）名称"          

            ]

VALUEARRAY = [[0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0],
             [0,0,0,0,0,0],
             [0,0,0,0,0],

             [0,0,0,0,0,0],
             [0,0,0,0,0,0],
             [0,0,0,0,0,0],
             [0,0,0,0,0,0],

             [0,0,0,0,0,0,0,0,0,0,0],
             [0,0,0,0,0,0],

             [0,0,0,0,0],
             [0],
             [0],
             [0]
             ]
            
ColumnARRAY=[]
count = 0
SchoolARRAY=[]
SchoolARRAYCount =[]
def initSchool():
    tmpdfs = pd.read_excel(excelroot)
    SchoolARRAY = list(tmpdfs['XXMC'])

def init():
     global ColumnARRAY
     for i  in range(1,14):
         mcolumn ="COLUMN_"+str(i)
         ColumnARRAY.append(mcolumn)
     for i  in range(13):
         print (ColumnARRAY[i])
def compute():
    global VALUEARRAY,count
    redfs = pd.DataFrame()

    for parent,dirnames,filenames in os.walk(dirroot): 
        for filename in filenames:
            file = dirroot + filename
            print (file)
            tmpdfs = pd.read_excel(file)
            dfs = tmpdfs.fillna("missing")
            print  (dfs)
            redfs = redfs.append(pd.DataFrame(dfs)) #写入文件
            frames=[pd.DataFrame(dfs['COLUMN_11']),pd.DataFrame(dfs['COLUMN_12']),pd.DataFrame(dfs['COLUMN_13'])]
            result = pd.concat(frames)
            result.to_excel(u'学校提取.xlsx')
                        
            for i in range(len(ColumnARRAY)-3):
    
                for j  in range(len(NAMEARRAY[i])):
                    if NAMEARRAY[i][j] == '其他'or i ==7:
                        rows=dfs[ColumnARRAY[i]].str.replace('<span>','').str.replace('</span>','').str.contains(NAMEARRAY[i][j])
                    else:
                        rows=dfs[ColumnARRAY[i]].str.replace('<span>','').str.replace('</span>','') ==NAMEARRAY[i][j]                       
                    VALUEARRAY[i][j] += dfs.loc[rows,ColumnARRAY[i]].count()
    redfs.to_excel(u'用人单位.xlsx')

def writeTXT():
    fileName = "datayrdw"+re.sub(r'[^0-9]','',str(datetime.datetime.now()))  +".csv"
    file_object  = open(fileName, 'w')
    try:
        for  i  in range(0,13):
            file_object.write(str(",")+str(QUEARRAY[i]))

            for j  in range(0,len(NAMEARRAY[i])):
                file_object.write(str(",")+NAMEARRAY[i][j])
            file_object.write("\r\n")
            
            file_object.write(str(","))

            for j  in range(0,len(VALUEARRAY[i])):
                file_object.write(str(",")+str(VALUEARRAY[i][j]))
            file_object.write("\r\n")

    finally:
        file_object.close( )
def main():
    init()
    initSchool()
    compute()
    writeTXT()
if __name__ =='__main__':
    main()
 
 


              
