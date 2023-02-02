import pymongo
from pymongo import MongoClient
from bson.objectid import ObjectId #這東西在透過ObjectID去尋找的時候會用到

import pandas as pd
import requests
import datetime

import numpy as np

from datetime import datetime
from datetime import timedelta
from datetime import date

import requests
import time

import os

import gspread
from oauth2client.service_account import ServiceAccountCredentials

import xlwt
import xlrd

import sys

# 這邊專門給我的 Module使用
from importlib import reload
# # 將project_home指定為你的專案路徑 要在同一個Disk裡面
# project_home = u'/Keep_the_desktop_clean/save_jupyter_python_code/ForXinxingProject/MyModule'
# if project_home not in sys.path:
#     sys.path = [project_home] + sys.path
# # 可以讀取整個工具組，也可以讀取特定function

import RJ
RJ = reload(RJ) # 因為我的module也一直在改，所以就...每次執行都要重新import
print(RJ.Owner)

# ================================================================================================================

# XinxingPG : 指的是專門給xinxing國中使用，作為畫圖與產生報表使用
Owner = "Ren Jhang, renxjhang@gmail.com - XinxingPG : 2022/03/07超大型更新excelFileAllData之後再加兩個欄位"

# ============================================= Grid_Xinxing ======================================================
# 一次所有報表全產生
# 舊版源自：http://localhost:8888/notebooks/ForXinxingProject/產生新興國中飲水量rank報表.ipynb  (有出Bug)
# 新版源自：http://localhost:8888/notebooks/ForXinxingProject/老師要我再多畫的圖/report_generator_校正.ipynb (2022/03/01更新)
# 更新版源自: http://localhost:8889/notebooks/ForXinxingProject/report_generator/最終確認.ipynb (2022/03/04更新, 超大改動, 不須依靠系統每天統計一次, 速度超大幅度提升)

def excelFileAllData(From, To, SavePath, Threshold=5):
    
    print("開始執行excelFileAllData...")
    print("-----> 開始讀取學生資料")
    # 取得所有身分是學生的學生名單, 並取出他們的學號, 名字要不等於unregistered --> 有學生嘗試鑽漏洞 2022/02/23

    search = {"$and": [{"Identity": "Student"}, 
                       {"Weight": {"$ne": None}}, 
                       {"Height": {"$ne": None}},
                       {"Weight": {"$ne": 0.0}}, 
                       {"Height": {"$ne": 0.0}}, 
                       {"Name": {"$ne": "unregistered"}}
                      ]}
    sort = [("SchoolID", 1)]
    Students = RJ.SC_ReadDataFromDB(DB='xinxing_dispenser', Collection='members_data', Search=search, Display={'_id':0}, Sort=sort)
    # print(Students)

    # print(len(Students))
    # 所有資料齊全的學生
    print("-----> 取用資料齊全的學生資訊")
    AllStudent = pd.DataFrame(columns = ['CardID', 'SchoolID', 'Name', 'Gender', 'Identity', 'Class', 'Grade', 'Room',
                                         'EnrollYear', 'Weight', 'Height', 'ShouldTake', 'BMI(value)', 'BMI(group)'])

    for i in range(len(Students)):
        BMI_value = Students[i]['Weight']/( (Students[i]['Height']/100)*(Students[i]['Height']/100) )
        BMI_value = np.round(BMI_value, 1)
        if np.isnan(BMI_value)!=True: # 有身高與體重的學生
            if BMI_value<18.5:
                BMI_group = "體重過輕"
            elif BMI_value>=18.5 and BMI_value<24:
                BMI_group = "體重正常"
            elif BMI_value>=24:
                BMI_group = "過胖族群"
            AppendData = {
                'CardID': Students[i]['CardID'],
                'SchoolID': Students[i]['SchoolID'], 
                'Name': Students[i]['Name'], 
                'BMI(value)': BMI_value,
                'BMI(group)': BMI_group,
                'Gender':Students[i]['Gender'],
                'ShouldTake':Students[i]['DailyShouldDrink'],
                'Identity':Students[i]['Identity'], 
                'EnrollYear':Students[i]['EnrollYear'], 
                'Class':Students[i]['Class'], 
                'Grade':Students[i]['Class'][:1], 
                'Room':Students[i]['Class'][1:], 
                'Weight':Students[i]['Weight'], 
                'Height':Students[i]['Height']
            }
        else:
            AppendData = {
                'CardID': Students[i]['CardID'],
                'SchoolID': Students[i]['SchoolID'], 
                'Name': Students[i]['Name'], 
                'BMI(value)': None,
                'BMI(group)': None,
                'Gender':Students[i]['Gender'],
                'ShouldTake':Students[i]['DailyShouldDrink'],
                'Identity':Students[i]['Identity'], 
                'EnrollYear':Students[i]['EnrollYear'], 
                'Class':Students[i]['Class'], 
                'Grade':Students[i]['Class'][:1], 
                'Room':Students[i]['Class'][1:], 
                'Weight':Students[i]['Weight'], 
                'Height':Students[i]['Height']            
            }        
        AllStudent = AllStudent.append(AppendData, ignore_index=True, sort=False)
    # print(AllStudent.iloc[0])

    # 抓raw_data塞成dataframe
    print("-----> 讀取原始數據(raw_data)")
    From = From + ' 00:00:00'
    To = To + ' 23:59:59'
    search = {'Timming':{'$gte': From, '$lte': To}, 'Dispenser':{'$regex':"^xinxing"}}
    sort = [("Timming",1)]
    ReadData = RJ.SC_ReadDataFromDB(DB='xinxing_dispenser', Collection='raw_data', Search=search, Display={"_id":0}, Sort=sort)
    # print(len(ReadData))

    AllRawData = pd.DataFrame(columns = ['Timming(Date)', 'Choose', 'WaterVolume', 'Dispenser', 'CardID', 'WaterTemp'])
    for i in range(len(ReadData)):
        WaterVolume = ReadData[i]['WaterVolume']
        # print("校正前", WaterVolume)
        # 校正WaterVolume
        if WaterVolume<0 :
            WaterVolume = float(int( WaterVolume + 1000000 ))
        if WaterVolume>10000 :
            WaterVolume = float(int( 1000000-WaterVolume ))
        AppendData = {
            'Timming(Date)': ReadData[i]['Timming'][:10],
            'Choose': ReadData[i]['Choose'], 
            'WaterVolume': WaterVolume, 
            'Dispenser': ReadData[i]['Dispenser'],
            'CardID': ReadData[i]['CardID'],
            'WaterTemp': ReadData[i]['WaterTemp']
        }
        AllRawData = AllRawData.append(AppendData, ignore_index=True, sort=False)

    # print(len(AllRawData))
    # print(AllRawData.iloc[0])

    # 完成整張New_AllRawData表: join
    print("-----> 將原始數據(raw_data)與學生資訊進行比對")
    New_AllRawData=AllRawData.join(AllStudent.set_index('CardID'), on='CardID',lsuffix='_l', rsuffix='_r')
    # print(New_AllRawData.iloc[0], "\n")
    # print(len(New_AllRawData))
    # print(New_AllRawData, "\n")

    # 取出目前SchoolID有資料的raw_data
    print("-----> 取出第一階段比對成功的數據(CardID:CardID)")
    OK_New_AllRawData = New_AllRawData[New_AllRawData.SchoolID.notnull()]
    # print(OK_New_AllRawData)
    print("-----> 第一階段有"+str(len(OK_New_AllRawData))+"筆比對成功的數據")
    # 取出目前SchoolID沒有資料的raw_data
    print("-----> 檢查第一階段是否有比對失敗的數據(CardID:CardID)")
    NaN_New_AllRawData = New_AllRawData[New_AllRawData.SchoolID.isnull()]
    # print(NaN_New_AllRawData)

    if NaN_New_AllRawData.empty == True:
        print("-----> 第一階段沒有比對失敗的數據")
        effectiveExcel = OK_New_AllRawData
    else:
        print("-----> 第一階段有"+str(len(NaN_New_AllRawData))+"筆比對失敗的數據")
        print("-----> 開始進一步處理第一階段比對失敗的數據")
        droplist = ['SchoolID', 'Name', 'Gender', 'Identity', 'Class', 'Grade', 'Room',
                    'EnrollYear', 'Weight', 'Height', 'ShouldTake', 'BMI(value)', 'BMI(group)']
        for i in range(len(droplist)):
            NaN_New_AllRawData = NaN_New_AllRawData.drop(columns=[droplist[i]])
        print("-----> 進行第二階段數據比對(CardID:SchoolID)")    
        Padding_New_AllRawData = NaN_New_AllRawData.join(AllStudent.set_index('SchoolID'), on='CardID',lsuffix='', rsuffix='_r')
        # print(Padding_New_AllRawData)
        print("-----> 剔除第二階段數據重複的欄位")
        Padding_New_AllRawData = Padding_New_AllRawData.drop(columns=['CardID_r'])
        #print(Padding_New_AllRawData)
        print("-----> 檢查第二階段是否有比對成功的數據")
        Padding_New_AllRawData = Padding_New_AllRawData[Padding_New_AllRawData.Name.notnull()]

        if Padding_New_AllRawData.empty == True:
            print("-----> 第二階段沒有比對成功的數據")
            effectiveExcel = OK_New_AllRawData
        else:
            print("-----> 第二階段有"+str(len(Padding_New_AllRawData))+"筆比對成功的數據")
            # 如果資料有對上，就把學號補上
            # print(Padding_New_AllRawData["CardID"])
            print("-----> 填補第二階段比對成功的數據")
            SchoolID_List = Padding_New_AllRawData["CardID"].tolist()
            # print(SchoolID_List)
            Padding_New_AllRawData["SchoolID"] = SchoolID_List
            # print(Padding_New_AllRawData)
            print("-----> 將 第一階段比對成功的數據 與 填補好的第二階段比對成功的數據 進行整合")
            # 合併兩個dataframe
            effectiveExcel = OK_New_AllRawData.append(Padding_New_AllRawData, ignore_index=True, sort=False)
            # print(effectiveExcel)


    effectiveExcel = effectiveExcel.reset_index(drop= True)
    print("-----> 開始進行數據整理(共"+str(len(effectiveExcel))+"筆有效資料)......")
    # print("effectiveExcel資料筆數:", len(effectiveExcel))
    # print(effectiveExcel)
    # effectiveExcel.to_excel('effectiveExcel.xlsx')  ################################ 輸出處理過的 raw_data excel 檔    


    # ##########################################################################################################    

    DateFrom = From[:10]
    DateTo = To[:10]
    DateFrom = RJ.AddOrMinusDays(DateFrom, delta=-1, display="%Y-%m-%d")

    Dates = []
    DateTakeWaterCnt = []
    UseOrNot = pd.DataFrame(columns = ['日期', '一天取水人數', '所設閥值', '採用當天資料與否'])

    AllDatesStatisticData = []

    while True:

        DateFrom = RJ.AddOrMinusDays(DateFrom, delta=1, display="%Y-%m-%d")
        print("-----> 計算"+str(DateFrom)+"當日各成員飲水總量等數據")

        cnt_num = 0 # 算當天有幾個人取水
        # 取得當天數據
        try: # 有當天數據
            DailyData = effectiveExcel.loc[effectiveExcel['Timming(Date)']==DateFrom]
            # 如果當日有數據，統計出當日所有學生的統計數據
            DailyStatisticData = []
            for i in range(len(AllStudent)):
                try: # raw_data有資料
                    #print("現在輪到", AllStudent.iloc[i]['SchoolID'])
                    DailyPeopleData = DailyData.loc[DailyData['SchoolID']==AllStudent.iloc[i]['SchoolID']]
                    # print(DailyPeopleData)
                    DailyPeopleSum = DailyPeopleData['WaterVolume'].sum()
                    # 產生單一個人當日的數據
                    try:
                        Percentage = float(np.round((DailyPeopleSum/DailyData.iloc[0]['ShouldTake'])*100, 2))
                    except:
                        Percentage = 0.0
                    DailyPeopleStatisticData = {
                                                '日期':DailyPeopleData.iloc[0]['Timming(Date)'],
                                                '年級':DailyPeopleData.iloc[0]['Grade'],
                                                '班級':DailyPeopleData.iloc[0]['Class'],
                                                '學號':DailyPeopleData.iloc[0]['SchoolID'],
                                                '姓名':DailyPeopleData.iloc[0]['Name'],
                                                '每日應喝水量(ml)':DailyPeopleData.iloc[0]['ShouldTake'],
                                                '本日飲水量':DailyPeopleSum,
                                                '本日飲水完成度(%)':Percentage
                                               }
                    cnt_num = cnt_num + 1
                except: # raw_data沒有資料
                    DailyPeopleStatisticData = {
                                                '日期':DailyData.iloc[0]['Timming(Date)'],
                                                '年級':AllStudent.iloc[i]['Grade'],
                                                '班級':AllStudent.iloc[i]['Class'],
                                                '學號':AllStudent.iloc[i]['SchoolID'],
                                                '姓名':AllStudent.iloc[i]['Name'],
                                                '每日應喝水量(ml)':AllStudent.iloc[i]['ShouldTake'],
                                                '本日飲水量':0.0,
                                                '本日飲水完成度(%)':0.0
                                               }        
                DailyStatisticData.append(DailyPeopleStatisticData)


        except: # 沒當天數據，所有學生都要歸0
            DailyStatisticData = []
            for i in range(len(AllStudent)):
                DailyPeopleStatisticData = {
                                            '日期':DateFrom,
                                            '年級':AllStudent.iloc[i]['Grade'],
                                            '班級':AllStudent.iloc[i]['Class'],
                                            '學號':AllStudent.iloc[i]['SchoolID'],
                                            '姓名':AllStudent.iloc[i]['Name'],
                                            '每日應喝水量(ml)':AllStudent.iloc[i]['ShouldTake'],
                                            '本日飲水量':0.0,
                                            '本日飲水完成度(%)':0.0
                                           }
                DailyStatisticData.append(DailyPeopleStatisticData)
        Dates.append(DateFrom)
        DateTakeWaterCnt.append(cnt_num)
        AllDatesStatisticData.append(DailyStatisticData)

        if DateFrom == DateTo: # 全算完就跳開迴圈
            break    

    print("-----> 顯示資料採計：")
    NotCountDates = []
    CountDates = []
    for i in range(len(Dates)):
        if DateTakeWaterCnt[i]<Threshold:
            NotCountDates.append(Dates[i])
            print("\t", Dates[i], DateTakeWaterCnt[i], "vs.", Threshold, "不採用")
            UAppendData = {
                '日期':Dates[i], 
                '一天取水人數': DateTakeWaterCnt[i],
                '所設閥值':Threshold,
                '採用當天資料與否': "No"
            }
            UseOrNot = UseOrNot.append(UAppendData, ignore_index=True, sort=False)

        else:
            CountDates.append(Dates[i])
            print("\t", Dates[i], DateTakeWaterCnt[i], "vs.", Threshold, "採用")
            UAppendData = {
                '日期':Dates[i], 
                '一天取水人數': DateTakeWaterCnt[i],
                '所設閥值':Threshold,
                '採用當天資料與否': "Yes"
            }
            UseOrNot = UseOrNot.append(UAppendData, ignore_index=True, sort=False)

    # ##########################################################################################################    

    ##################################################### 如果CountDates==0 --> 請重新定義Thredshold
    LackData = pd.DataFrame(columns = ['CardID', 'SchoolID', 'Name', 'Gender', 'Identity', 'Class', 'Grade', 'Room',
                                         'EnrollYear', 'Weight', 'Height', 'ShouldTake', 'BMI(value)', 'BMI(group)'])
    if CountDates!=[]:
        print("-----> 進行此期間資料統計的計算平均(生成最終依據表單)")
        # 把每天同一個人的資料抓出來進行統計
        # 要抓取哪些第幾天的資料
        whichIndexList = []
        for i in range(len(CountDates)):
            whichIndex = Dates.index(CountDates[i])
            whichIndexList.append(whichIndex)

        # 對每個人進行最終的統計Loop
        Export_Excel = pd.DataFrame(columns = ['年級', '班級', '學號', '姓名', 
                                               '期間平均每日飲水完成度(%)', '期間平均每日飲水量', '每日應喝水量(ml)', '當前紀錄體重(kg)', '每單位體重建議飲水量(ml)'])
        for i in range(len(AllStudent)):
            Average_Take = 0.0
            Average_Percent = 0.0
            TotalTake = 0.0
            for j in whichIndexList :
                TotalTake = TotalTake + AllDatesStatisticData[j][i]['本日飲水量']
            Average_Take = float(np.round(TotalTake/len(whichIndexList), 2))
            try:
                Average_Percent = float(np.round((Average_Take/AllStudent.iloc[i]['ShouldTake'])*100, 2))
            except:
                Average_Percent = 0

            AppendData = {
                '年級': AllDatesStatisticData[0][i]['年級'], 
                '班級': AllDatesStatisticData[0][i]['班級'], 
                '學號': AllDatesStatisticData[0][i]['學號'], 
                '姓名': AllDatesStatisticData[0][i]['姓名'], 
                '期間平均每日飲水完成度(%)':Average_Percent,
                '期間平均每日飲水量': Average_Take, 
                '每日應喝水量(ml)': AllDatesStatisticData[0][i]['每日應喝水量(ml)'],
                '當前紀錄體重(kg)':AllStudent.iloc[i]['Weight'],
                '每單位體重建議飲水量(ml)':40.0
            }
            CheckAppendData = pd.DataFrame.from_dict(AppendData, orient='index')
            CheckAppendData = np.any(CheckAppendData.isnull())        
            if CheckAppendData==True:
                # print("-----> 注意：有資料不完整的學生")
                LackData = LackData.append(AllStudent.iloc[i].to_dict(), ignore_index=True, sort=False)
                CheckAppendData=False
            else:
                Export_Excel = Export_Excel.append(AppendData, ignore_index=True, sort=False)

            #### 正常執行的
    else:
        print("-----> 無法生成報表，請重新定義Thredshold")

    if len(LackData)!=0:  
        print("-----> 注意：有資料不完整的學生，名單將會呈現在產生的報表首頁")
        LackData = LackData.drop(columns=['EnrollYear'])
        LackData = LackData.drop(columns=['ShouldTake'])
        LackData = LackData.drop(columns=['BMI(value)'])
        LackData = LackData.drop(columns=['BMI(group)'])

    # ##########################################################################################################     

    # 生成報表
    # 輸出成excel表
    # Create a Pandas Excel writer using XlsxWriter as the engine.
    DateFrom = From[:10]
    DateTo = To[:10]
    fileName = SavePath+'新興國中飲水Ranking('+DateFrom+'到'+DateTo+').xlsx'
    writer = pd.ExcelWriter(fileName, engine='xlsxwriter')

    if len(LackData)!=0: 
        print("-----> 開始生成報表: 欠缺資料的學生")
        # Convert the dataframe to an XlsxWriter Excel object.
        LackData.to_excel(writer, sheet_name='欠缺資料的學生')

    print("-----> 開始生成報表: 資料採計明細")
    # Convert the dataframe to an XlsxWriter Excel object.
    UseOrNot.to_excel(writer, sheet_name='資料採計明細')

    print("-----> ! 注意: 報表中 '期間平均每日飲水完成度(%)' 若超過100僅會以100顯示 !")

    print("-----> 開始生成報表: 全校(個人)")
    SchoolRank_Person = Export_Excel.sort_values(by=['期間平均每日飲水完成度(%)'], ascending=False)
    SchoolRank_Person = SchoolRank_Person.reset_index(drop= True)
    # print(SchoolRank_Person)
    SchoolRank_Person.loc[SchoolRank_Person['期間平均每日飲水完成度(%)']>100.0, '期間平均每日飲水完成度(%)']=100.0
    # Convert the dataframe to an XlsxWriter Excel object.
    SchoolRank_Person.to_excel(writer, sheet_name='全校(個人)')

    print("-----> 開始生成報表: 各年級(個人)")
    # 找出有幾種類別
    Grades = SchoolRank_Person['年級'].unique()
    GradeRank_Person = []
    for grade in Grades:
        gradeRank_Person = SchoolRank_Person[SchoolRank_Person['年級']==grade]    
        gradeRank_Person = gradeRank_Person.sort_values(by=['期間平均每日飲水完成度(%)'], ascending=False)
        gradeRank_Person = gradeRank_Person.reset_index(drop= True)
        gradeRank_Person.to_excel(writer, sheet_name=str(grade)+'年級(個人)')
        GradeRank_Person.append(gradeRank_Person)

    print("-----> 開始生成報表: 各班級(個人)")
    # 也要順便生成以班級為單位的表
    SchoolRank_Class = pd.DataFrame(columns = ['年級', '班級', '班級人數(排除資料欠缺者)', '期間平均每日飲水完成度(%)'])
    for i in range(len(GradeRank_Person)):
        gradeRank_Person = GradeRank_Person[i]
        Classes = gradeRank_Person['班級'].unique()
        for Class in Classes:
            ClassRank_Person = gradeRank_Person[gradeRank_Person['班級']==Class]
            ClassRank_Person = ClassRank_Person.sort_values(by=['期間平均每日飲水完成度(%)'], ascending=False)
            ClassRank_Person = ClassRank_Person.reset_index(drop= True)        
            ClassRank_Person.to_excel(writer, sheet_name=str(Class)+'(個人)')
            # 統計這個班的總和
            AppendData = {
                '年級': ClassRank_Person.iloc[0]['年級'],
                '班級': ClassRank_Person.iloc[0]['班級'], 
                '班級人數(排除資料欠缺者)':len(ClassRank_Person), 
                '期間平均每日飲水完成度(%)': float(np.round(ClassRank_Person['期間平均每日飲水完成度(%)'].sum()/len(ClassRank_Person), 2)), 
            }
            SchoolRank_Class = SchoolRank_Class.append(AppendData, ignore_index=True, sort=False)

    print("-----> 開始生成報表: 全校(班級)")
    SchoolRank_Class = SchoolRank_Class.sort_values(by=['期間平均每日飲水完成度(%)'], ascending=False)
    SchoolRank_Class = SchoolRank_Class.reset_index(drop= True)
    SchoolRank_Class.to_excel(writer, sheet_name='全校(班級)')

    print("-----> 開始生成報表: 各年級(班級)")
    Grades2 = SchoolRank_Class['年級'].unique()
    # print(SchoolRank_Class[SchoolRank_Class['年級']==100.0])
    # print("Grades2", Grades2)
    for grade in Grades2:
        GradeRank_Class = SchoolRank_Class[SchoolRank_Class['年級']==grade]
        GradeRank_Class.sort_values(by=['期間平均每日飲水完成度(%)'], ascending=False)
        GradeRank_Class = GradeRank_Class.reset_index(drop= True)
        GradeRank_Class.to_excel(writer, sheet_name=grade+'年級(班級)')

    # Close the Pandas Excel writer and output the Excel file.
    writer.save()
    print("-----> 報表生成完畢")
    print("\n===========================\n\n")    
