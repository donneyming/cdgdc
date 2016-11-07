# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""

import pandas as pd
import numpy as np
import re 
import time
import os
import os.path
import datetime

dirroot = "C:\\Users\\cms\\Desktop\\\学科评估\\问卷调查\\回收\\问卷【研究生导师指导情况调查问卷】调查结果明细表\\"




NAMEARRAY = [["男","女"],
            ["博士研究生导师","硕士研究生导师"],
            ["35岁及以下","36-45岁","46-55岁","56-60岁","61岁及以上"],
            ["正高级（教授、研究员等）","副高级（副教授、副研究员等）","中级（讲师、助理研究员等）","其他"],
            ["校级领导","院（系）级领导","未担任行政职务","其他"],
            ["0","1至4","5","6","7","8","9","10","11","12","13","14","15","16","17","18","19","20人及以上"],
            ["0","1至4","5","6","7","8","9","10","11","12","13","14","15","16","17","18","19","20人及以上"],
            ["良师益友型","普通师生型","松散疏离型","老板员工型"],
            ["学术规范","学术研究","学业内容","职业规划","日常生活","其他"],##单独处理
            ["当面一对一交流","组会（如课题组会、学术沙龙等）","语音交流（如电话、网络语音等）","文字交流（邮件、信件短信等）","其他"],##单独处理
            ["几乎不见面","1","2","3","4","5","6","7","8","9","10","11","12","13","14","15及以上"],
            ["0","1","2","3","4","5","6","7","8","9","10及以上"],
            ["0","1","2","3","4","5","6","7","8","9","10及以上"],
            ["较大，是主要完成人之一","一般，仅仅参与部分工作","较小，参与一些辅助工作","其他"],##单独处理
            ["不相关","不太相关","一般","比较相关","很相关"],
            ["1","2","3","4","5","6"],
            ["1","2","3","4","5","6"],
            ["1","2","3","4","5","6"],
            ["1","2","3","4","5","6"],
            ["1","2","3","4","5","6"],
            ["1","2","3","4","5","6"],
            ["1","2","3","4","5","6"],
            ["1","2","3","4","5","6"],
            ["1","2","3","4","5","6"],
            ["自我学习能力","学术研究能力","独立思考能力","创新能力","沟通表达能力","组织协调能力","团队协作能力","心理调适能力","实践动手能力","其他"],##单独处理
            ["很不满意","不满意","不太满意","基本满意","满意","很满意"],
            ["是","否"],
            ["是","否"],
            ["没有直接指导","偶尔指导","与实际指导教师的指导程度相当"]
            ]            
            
QUEARRAY = ["您导师的性别是",
            "您导师的类型是",
            "您导师的年龄是",
            "您导师的职称是",
            "您导师担任的行政职务是",
            "目前，您导师指导的在读博士生人数约为",
            "在读硕士生人数约为",
            "您认为与导师之间的关系是",
            "您日常与导师交流的主要内容是：（最多选3项）",
            "导师进行学术指导的主要方式是：（最多选3项）",
            "导师平均每月面对面（含视频）对您进行学术指导的有效时间为小时",
            "您在攻读当前学位层次期间参与课题研究的总数是",
            "其中参与的导师课题数是",
            "在您参与度最高的导师课题中，您所发挥的作用是",
            "您认为参与的导师课题与您学位论文的关系是：",
#            "请对导师在以下方面对您进行指导的满意度进行评分",
            "（1）人生观、价值观的养成",
            "（2）学术道德与规范培养",
            "（3）课程学习",
            "（4）创新能力培养",
            "（5）实践应用能力培养",
            "（6）研究方法训练",
            "（7）学位论文撰写",
            "（8）解决日常生活困难",
            "（9）引导职业生涯规划",
            "通过导师的指导培养，您觉得自己哪些方面的能力得到了提升（（最多选5项）",
            "您对导师指导质量的总体满意程度是：",
            "您的导师与您在学校登记的指导教师是否为同一人？",
            "您在填答此问卷时，是否受到学校有意引导或干扰？",
            "您在学校登记的指导教师对您的指导情况是"
            ]

VALUEARRAY = [[0,0],#男教师数量#女教师数量
#博士研究生导师数目
#硕士研究生导师数目
             [0,0],
             [0,0,0,0,0],
             [0,0,0,0],
             [0,0,0,0],
             [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0],
             [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0],
             [0,0,0,0],
             [0,0,0,0,0,0],
             [0,0,0,0,0],
             [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0],
             [0,0,0,0,0,0,0,0,0,0,0],
             [0,0,0,0,0,0,0,0,0,0,0],

             [0,0,0,0],
             [0,0,0,0,0],

             [0,0,0,0,0,0],
             [0,0,0,0,0,0],
             [0,0,0,0,0,0],
             [0,0,0,0,0,0],
             [0,0,0,0,0,0],
             [0,0,0,0,0,0],
             [0,0,0,0,0,0],
             [0,0,0,0,0,0],
             [0,0,0,0,0,0],

             [0,0,0,0,0,0,0,0,0,0],
             [0,0,0,0,0,0],

             [0,0],
             [0,0],
             [0,0,0]
             ]
            
ColumnARRAY=[]
count = 0

def process(value,typeOfValue):
    result =""
    if typeOfValue == 2:
        result = value.replace('<span>','')
        result = value.replace('</span>','')
        print ("value is %s"%result)
    return result
def init():
     global ColumnARRAY
     for i  in range(1,30):
         mcolumn ="COLUMN_"+str(i)
         ColumnARRAY.append(mcolumn)
#         print (mcolumn)
def compute():
    global VALUEARRAY,count
    redfs = pd.DataFrame()
    print ("compute")
    for parent,dirnames,filenames in os.walk(dirroot): 
        for filename in filenames:
            file = dirroot + filename
            print (file)
            tmpdfs = pd.read_excel(file)
            dfs = tmpdfs.fillna("missing")
            redfs = redfs.append(pd.DataFrame(dfs)) #写入文件

            for i in range(29):
                #print("i is %d"%(len(NAMEARRAY[i])))

                for j  in range(0,len(NAMEARRAY[i])):
                      if NAMEARRAY[i][j] == '其他'or i ==8 or i ==9  or i ==14 or i ==24:
                          rows=dfs[ColumnARRAY[i]].str.replace('<span>','').str.replace('</span>','').str.contains(NAMEARRAY[i][j])
                      else:
                          rows=dfs[ColumnARRAY[i]].str.replace('<span>','').str.replace('</span>','') ==(NAMEARRAY[i][j])
                      VALUEARRAY[i][j] += dfs.loc[rows,ColumnARRAY[i]].count()
#    fileName = u"在校生汇总"+re.sub(r'[^0-9]','',str(datetime.datetime.now()))  +".xlsx"  
    redfs.to_excel(u"在校生汇总.xlsx")
    print ("compute finish")


def writeTXT():
    fileName = "datazxs"+re.sub(r'[^0-9]','',str(datetime.datetime.now()))  +".csv"
    file_object  = open(fileName, 'w')
    try:
        for  i  in range(0,29):
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
    compute()
    writeTXT()
if __name__ =='__main__':
    main()
 
 


              
