#!/usr/bin/env python
# -*- coding: UTF-8 -*-
'''=================================================
@Project -> File   ：TestMyExcel.py -> DealNeiTuiData
@Author ：Mr. Young
@Date   ：2021/4/14 14:26
@Desc   ：对内退及其他部分进行统计分析
=================================================='''

import pandas as pd
import numpy as np
import math
ExcelName = './样本.xls'

Data_Project_Sheet1 = pd.read_excel(ExcelName,sheet_name='内退及其他（24）')
data_project = Data_Project_Sheet1.values
data_project = data_project[1:,:]
# print("结果为:\n{0}".format(data_project))
#对项目所需要的数据进行分析

#1：首先把所有部门提取出来
projectNameList = sorted(list(set(data_project[:,1])))
projectNameList = np.array(projectNameList)

#：统计人数所需必要函数
# 判断部门领导、部门副职、员工
LingDaoAll = ['党委副书记/总经理','总会计师','副总经理','总工程师、副总经理','总经济师、副总经理','纪委副书记',
              '安全总监','两办副主任（保留正职职级）','主任','部长','副部长（保留正职职级）']
FuZhiAll = ['副部长','副主任','副部长（主持工作）','副主任（主持工作）','副总经理（中层副职待遇）','团委书记']
def defineCarrer(data):
    LingDaoCount = 0
    FuZhiCount = 0
    YuanGongCount = 0
    # if data in LingDaoAll:
    #     LingDaoCount = LingDaoCount + 1
    # elif data in FuZhiAll:
    #     FuZhiCount = FuZhiCount + 1
    # else:
    #     YuanGongCount = YuanGongCount + 1
    return LingDaoCount,FuZhiCount,YuanGongCount

# 判断职称 教高、高级、中级、初级、初级以下
JiaoGaoAll = ['教授级高工']
GaoJiAll = ['高级工程师','高级经济师','高级会计师','高级政工师']
ZhongJiAll = ['工程师','会计师','经济师','政工师','中级','技师']
ChuJiAll = ['助理工程师','助理会计师','助理经济师','助理政工师','技术员','会计员','政工员']
def defineLevel(data1):
    JiaoGaoCount = 0
    GaoJiCount = 0
    ZhongJiCount = 0
    ChuJiCount = 0
    ChuJiBelowCount = 0
    if data1 in JiaoGaoAll:
        JiaoGaoCount = JiaoGaoCount + 1
    elif data1 in GaoJiAll:
        GaoJiCount = GaoJiCount + 1
    elif data1 in ZhongJiAll:
        ZhongJiCount = ZhongJiCount + 1
    elif data1 in ChuJiAll:
        ChuJiCount = ChuJiCount + 1
    else:
        ChuJiBelowCount = ChuJiBelowCount + 1
    return JiaoGaoCount,GaoJiCount,ZhongJiCount,ChuJiCount,ChuJiBelowCount

# 学历统计 研究生、本科、专科、中专、中专及以下 按照最高学历
def defineStudyLevel(data_level1,data_level2):
    YanJiuShengCount = 0
    BenKeCount = 0
    ZhuanKeCount = 0
    ZhongZhuanCount = 0
    ZhongZhuanBelowCount = 0
    if isinstance(data_level2,str):
        if data_level2 == '研究生' or data_level2 =='硕士' or data_level2 =='硕士学位' or data_level2 =='研究生（函授）':
            YanJiuShengCount = YanJiuShengCount + 1
        elif data_level2 == '本科':
            BenKeCount = BenKeCount + 1
        elif data_level2 == '大专':
            ZhuanKeCount = ZhuanKeCount + 1
        elif data_level2 == '中专':
            ZhongZhuanCount = ZhongZhuanCount + 1
        else:
            ZhongZhuanBelowCount = ZhongZhuanBelowCount + 1
    else:
        if data_level1 == '研究生':
            YanJiuShengCount = YanJiuShengCount + 1
        elif data_level1 == '本科':
            BenKeCount = BenKeCount + 1
        elif data_level1 == '大专':
            ZhuanKeCount = ZhuanKeCount + 1
        elif data_level1 == '中专':
            ZhongZhuanCount = ZhongZhuanCount + 1
        else:
            ZhongZhuanBelowCount = ZhongZhuanBelowCount + 1

    return YanJiuShengCount,BenKeCount,ZhuanKeCount,ZhongZhuanCount,ZhongZhuanBelowCount

#2:通过遍历上述部门 获取需要统计的数据
#数据写入文件 ---->项目人员统计  NeiTuiResults.txt
fileName_Project =  'NeiTuiResults.txt'
projectTitle = ['人数','部门领导','部门副职','员工','男','女','教高','高级','中级','初级','初级以下','研究生','本科',
                   '专科','中专','中专及以下']
with open(fileName_Project,'w') as project:
    for title in projectTitle:
        project.write(title + '\t')
    project.write('\n')

#对各项数据总数进行统计
PeopleNumsSum = 0
LingDaoSum = 0
FuZhiSum = 0
YuanGongSum = 0

MaleSum = 0
FmaleSum = 0

JiaoGaoSum = 0
GaoJiSum = 0
ZhongJiSum = 0
ChuJiSum = 0
ChuJi_Below_Sum = 0

YanJiuShengSum = 0
BenKeSum = 0
ZhuanKeSum = 0
ZhongZhuanSum = 0
ZhongZhuan_Below_Sum = 0

for pName in projectNameList:
    PeopleNums = 0
    LingDao = 0  #*******
    FuZhi = 0
    YuanGong = 0
    # 性别
    Male = 0  # 女
    Fmale = 0
    # 职称
    JiaoGao = 0
    Gaoji = 0
    ZhongJi = 0
    ChuJi = 0
    ChuJi_Below = 0

    # 学历
    YanJiuSheng = 0
    BenKe = 0
    ZhuanKe = 0  #**********
    ZhongZhuan = 0
    ZhongZhuan_Below = 0
    #开始遍历原始数据进行统计
    for dataItem in data_project:
        if pName == dataItem[1]:
            PeopleNums = PeopleNums + 1  #人数
            [LingDaoCount,FuZhiCount,YuanGongCount] = defineCarrer(dataItem[14]) #职级
            LingDao = LingDao + LingDaoCount
            FuZhi = FuZhi + FuZhiCount
            YuanGong = YuanGong + YuanGongCount

            #性别
            if dataItem[4] == '男':
                Fmale = Fmale + 1
            else:
                Male = Male + 1

            #职称评定
            [JiaoGaoCount,GaoJiCount,ZhongJiCount,ChuJiCount,ChuJiBelowCount] = defineLevel(dataItem[14])
            JiaoGao = JiaoGao + JiaoGaoCount
            Gaoji = Gaoji + GaoJiCount
            ZhongJi = ZhongJi + ZhongJiCount
            ChuJi = ChuJi + ChuJiCount
            ChuJi_Below = ChuJi_Below + ChuJiBelowCount
            #学历评定
            print(pName+'\n')
            print(dataItem[7],dataItem[10])
            [YanJiuShengCount,BenKeCount,ZhuanKeCount,ZhongZhuanCount,ZhongZhuanBelowCount] = defineStudyLevel(dataItem[7],dataItem[10])
            YanJiuSheng = YanJiuSheng + YanJiuShengCount
            BenKe = BenKe + BenKeCount
            ZhuanKe = ZhuanKe + ZhuanKeCount
            ZhongZhuan = ZhongZhuan + ZhongZhuanCount
            ZhongZhuan_Below = ZhongZhuan_Below + ZhongZhuanBelowCount

        projectResults = [PeopleNums,LingDao,FuZhi,YuanGong,Male,Fmale,JiaoGao,Gaoji,ZhongJi,ChuJi,ChuJi_Below,YanJiuSheng,BenKe,
                          ZhuanKe,ZhongZhuan,ZhongZhuan_Below]

    PeopleNumsSum = PeopleNumsSum + PeopleNums
    LingDaoSum = LingDaoSum + LingDao
    FuZhiSum = FuZhiSum + FuZhi
    YuanGongSum = YuanGongSum + YuanGong
    MaleSum = MaleSum + Male
    FmaleSum = FmaleSum + Fmale
    JiaoGaoSum = JiaoGaoSum + JiaoGao
    GaoJiSum = GaoJiSum + Gaoji
    ZhongJiSum = ZhongJiSum + ZhongJi
    ChuJiSum = ChuJiSum + ChuJi
    ChuJi_Below_Sum = ChuJi_Below_Sum + ChuJi_Below
    YanJiuShengSum = YanJiuShengSum + YanJiuSheng
    BenKeSum = BenKeSum + BenKe
    ZhuanKeSum = ZhuanKeSum + ZhuanKe
    ZhongZhuanSum = ZhongZhuanSum + ZhongZhuan
    ZhongZhuan_Below_Sum = ZhongZhuan_Below_Sum + ZhongZhuan_Below

    with open(fileName_Project,'a') as project1:
        project1.write(pName+'\n')
        for projectdata in projectResults:
            project1.write(str(projectdata) + '\t')
        project1.write('\n')

SumResults = [PeopleNumsSum,LingDaoSum,FuZhiSum,YuanGongSum,MaleSum,FmaleSum,JiaoGaoSum,GaoJiSum,ZhongJiSum,
              ChuJiSum,ChuJi_Below_Sum,YanJiuShengSum,BenKeSum,ZhuanKeSum,ZhongZhuanSum,ZhongZhuan_Below_Sum]
with open(fileName_Project,'a') as project2:
    project2.write('各项总数为：'+ '\n')
    for sumdata in SumResults:
        project2.write(str(sumdata) + '\t')
    project2.write('\n')