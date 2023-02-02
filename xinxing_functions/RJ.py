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
import json

# import gspread
# from oauth2client.service_account import ServiceAccountCredentials

################### # 加入自製Module
# import sys
# from importlib import reload
# # 將project_home指定為你的專案路徑 要在同一個Disk裡面
# project_home = u'/Keep_the_desktop_clean/save_jupyter_python_code/ForXinxingProject/MyModule'
# if project_home not in sys.path:
#     sys.path = [project_home] + sys.path
# # 可以讀取整個工具組，也可以讀取特定function
# import RJ
# RJ = reload(RJ) # 因為我的module也一直在改，所以就...每次執行都要重新import
# print(RJ.Owner)


Owner = "Ren Jhang, renxjhang@gmail.com - 更新: 2022/05/26"

mongo_url_01 = "mongodb://admin:bmwee8097218@140.118.122.115:30415/"
mongo_url_02 = "mongodb://admin:bmwee8097218@140.118.122.115:30415/"

# 把藉由SC_ReadDataFromDB從資料庫讀取回來的資料從list轉型成json或dataframe
def convertListToAnotherForm(ReadData, convertTo="dataframe"):
    if isinstance(ReadData, list)!=True:
        return False
    if convertTo=="json":
        jsonReadData = json.dumps(ReadData)
        return jsonReadData
    elif convertTo=="dataframe":
        jsonReadData = json.dumps(ReadData)
        dfReadData = pd.read_json(jsonReadData)        
        return dfReadData
    else:
        return False

# ======================================================================================================== MongoDB SC Database 

############################# 目前功能有：
# 01. 讀取資料庫的資料
# 02. 做DB aggregate
# 03. 為資料庫現下的每一筆資料新增一個欄位
# 04. 更新資料到DB中的Collection: smartcampus的, 只更新一筆資料
# 05. 新增資料到DB中的Collection: smartcampus的, 只新增一筆資料
# 06. 更新或新增資料到DB中的Collection: smartcampus的, 只更新或新增一筆資料
# 07. 刪除一筆或所有符合條件的資料，慎用!!!!!!!!!!!!!

# 讀取資料庫的資料
def SC_ReadDataFromDB(DB, Collection, Search={}, Display={}, Sort=[], Limit=False):
    global mongo_url_01, mongo_url_02
    # input範例
    # DB = 'xinxing_dispenser'
    # Collection = 'raw_data'
    # Search={'Timming':{'$gte': From, '$lte': To}, 'Dispenser':{'$regex':"^xinxing"}, 'CardID':CardID}
    # 比較複雜的Search
    # Search = {"Date":Date, 'Name':{'$not':re.compile("^any.*")}, '$and':[{"Percent":{'$ne':"nan%"}}, {"Percent":{'$ne':"0.0%"}}]}
    # Display={"CardID":1} # 只要某個欄位，但"_id"仍然會顯示
    # Display={"_id":0} # 不顯示"_id"
    # Sort=[("CardID",1)]
    # Limit=False # 不限制抓回來的資料筆數 (有多少抓多少)
    # Limit=1 # 限制抓回來的資料筆數 (只抓一筆)
    # 正則表達式更多範例:
    # search = {"sort":"Program", "key":{'$regex':"^DS:A.*-ProgramStatus$"}} 
    # ^DS:A --> 以「DS:A」開頭 
    # .*    --> 這個部分隨便
    # -ProgramStatus$ --> 以「-ProgramStatus」結尾
    try:
        conn = MongoClient(mongo_url_01) 
        db = conn[DB]
        collection = db[Collection]
        if Sort!=[]:
            if Display == {}: #cursor裡面不要放Display不然會變成空白                
                if Limit==False:
                    cursor = collection.find(Search).sort(Sort)
                else:
                    cursor = collection.find(Search).sort(Sort).limit(Limit)
            else:
                if Limit==False:
                    cursor = collection.find(Search, Display).sort(Sort)
                else:
                    cursor = collection.find(Search, Display).sort(Sort).limit(Limit)
        else:
            if Display == {}: #cursor裡面不要放Display不然會變成空白
                if Limit==False:
                    cursor = collection.find(Search)
                else:
                    cursor = collection.find(Search).limit(Limit)
            else:
                if Limit==False:
                    cursor = collection.find(Search, Display)
                else:
                    cursor = collection.find(Search, Display).limit(Limit)
                
    except:
        conn = MongoClient(mongo_url_02) 
        db = conn[DB]
        collection = db[Collection]
        if Sort!=[]:
            if Display == {}: #cursor裡面不要放Display不然會變成空白                
                if Limit==False:
                    cursor = collection.find(Search).sort(Sort)
                else:
                    cursor = collection.find(Search).sort(Sort).limit(Limit)
            else:
                if Limit==False:
                    cursor = collection.find(Search, Display).sort(Sort)
                else:
                    cursor = collection.find(Search, Display).sort(Sort).limit(Limit)
        else:
            if Display == {}: #cursor裡面不要放Display不然會變成空白
                if Limit==False:
                    cursor = collection.find(Search)
                else:
                    cursor = collection.find(Search).limit(Limit)
            else:
                if Limit==False:
                    cursor = collection.find(Search, Display)
                else:
                    cursor = collection.find(Search, Display).limit(Limit)         
            
    #此處須注意，其回傳的並不是資料本身，你必須在迴圈中逐一讀出來的過程中，它才真的會去資料庫把資料撈出來給你。
    data = [d for d in cursor] #這樣才能真正從資料庫把資料庫撈到python的暫存記憶體中。
    if data==[]:
        return False
    else:
        return data

# 做DB aggregate
def SC_Aggregate(DB, Collection, Match, GroupAndGenerate):
    global mongo_url_01, mongo_url_02
    # 使用範例
    # Match = {"$match": {"DS_num":"test001"}}
    # GroupAndGenerate = {"$group": {"_id": {"DS_num":"$DS_num"}, "TotalHour":{"$sum": 1}, "TotalPress": {"$sum": '$Press'}, "TotalWater":{"$sum":"$Water"}}}
    try:
        conn = MongoClient(mongo_url_01) 
        db = conn[DB]
        collection = db[Collection]
        cursor = collection.aggregate([ Match, GroupAndGenerate ])
    except:
        conn = MongoClient(mongo_url_02) 
        db = conn[DB]
        collection = db[Collection]
        cursor = collection.aggregate([ Match, GroupAndGenerate ])

    
    data = [d for d in cursor] #這樣才能真正從資料庫把資料庫撈到python的暫存記憶體中。
    if data==[]:
        return False
    else:
        return data

# 為資料庫現下的每一筆資料新增一個欄位
def SC_AddNewOneColumn(DB, Collection, NewColumnName, DefaultValue="None"):
    global mongo_url_01, mongo_url_02
    ReadData = SC_ReadDataFromDB(DB=DB, Collection=Collection, Search={}, Display={"_id":1}, Sort=[])
    for i in range(len(ReadData)):
        DataID = ReadData[i]['_id']
        print(DataID)
        Search={"_id":DataID}
        UploadContent = {NewColumnName:DefaultValue}
        SC_UpdateInsertToDB(DB=DB, Collection=Collection, Search=Search, UploadContent=UploadContent)

# 更新資料到DB中的Collection: smartcampus的, 只更新一筆資料
def SC_UpdateToDB(DB, Collection, Search, UploadContent):
    global mongo_url_01, mongo_url_02
    try:
        conn = MongoClient(mongo_url_01) 
        db = conn[DB]
        collection = db[Collection]
        cursor = collection.update_one(Search, {'$set': UploadContent}, upsert=False)
    except:
        conn = MongoClient(mongo_url_02) 
        db = conn[DB]
        collection = db[Collection]
        cursor = collection.update_one(Search, {'$set': UploadContent}, upsert=False)

# 新增資料到DB中的Collection: smartcampus的, 只新增一筆資料
def SC_InsertToDB(DB, Collection, InsertContent):
    global mongo_url_01, mongo_url_02
    try:
        conn = MongoClient(mongo_url_01) 
        db = conn[DB]
        collection = db[Collection]
        cursor = collection.insert_one(InsertContent)
    except:
        conn = MongoClient(mongo_url_02) 
        db = conn[DB]
        collection = db[Collection]
        cursor = collection.insert_one(InsertContent)
            
            
# 更新或新增資料到DB中的Collection: smartcampus的, 只更新一筆資料
def SC_UpdateInsertToDB(DB, Collection, Search, UploadContent):
    global mongo_url_01, mongo_url_02
    try:
        conn = MongoClient(mongo_url_01) 
        db = conn[DB]
        collection = db[Collection]
        cursor = collection.update_one(Search, {'$set': UploadContent}, upsert=True)
    except:
        conn = MongoClient(mongo_url_02) 
        db = conn[DB]
        collection = db[Collection]
        cursor = collection.update_one(Search, {'$set': UploadContent}, upsert=True)

# 刪除一筆或所有符合條件的資料，慎用!!!!!!!!!!!!!
def SC_DeleteDataFromDB(DB, Collection, Condition, Delete="One"):
    global mongo_url_01, mongo_url_02
    try:
        conn = MongoClient(mongo_url_01) 
        db = conn[DB]
        collection = db[Collection]
        if Delete=="One":
            cursor = collection.delete_one(Condition)
        elif Delete=="Many":
            cursor = collection.delete_many(Condition)        
    except:
        conn = MongoClient(mongo_url_02) 
        db = conn[DB]
        collection = db[Collection]
        if Delete=="One":
            cursor = collection.delete_one(Condition)
        elif Delete=="Many":
            cursor = collection.delete_many(Condition) 
          

# ======================================================================================================================= About Time

def Get_now_dateAndTime(display="%Y_%m_%d %H_%M_%S"):
    nowTime = int(time.time()) # 取得現在時間
    struct_time = time.localtime(nowTime) # 轉換成時間元組
    timeString = time.strftime(display, struct_time) # 將時間元組轉換成想要的字串
    return timeString

def Get_date_of_today(display="%Y-%m-%d"):
    today = date.today()
    str_today = today.strftime(display)
    return str_today

def Get_date_of_tomorrow(display="%Y-%m-%d"):
    today = date.today()
    tomorrow =  today + timedelta(days=1)
    str_tomorrow = tomorrow.strftime(display)
    return str_tomorrow

def Get_date_of_yesterday(display="%Y-%m-%d"):
    today = date.today()
    yesterday =  today- timedelta(days=1)
    str_yesterday = yesterday.strftime(display)
    return str_yesterday

# 加/減幾天
def AddOrMinusDays(str_timming, delta=1, display="%Y-%m-%d"):
    t_timming = datetime.strptime(str_timming, display) # 轉成<class 'datetime.datetime'>
    # print(t_timming)
    t_timming = t_timming + timedelta(days=delta)
    str_new_timming = t_timming.strftime(display)
    # print(str_new_timming)
    return str_new_timming

# 算兩個時間的時間差(回傳字串訊息)
def timeDiff_msg(time_1="", time_2=""):
    # 抓現下時間
    display = "%Y-%m-%d %H:%M:%S"
    now = time.time()
    now = datetime.fromtimestamp(now)
    str_now = now.strftime(display) 
    if time_1=="" :
        time_1 = str_now
    if time_2=="" :
        time_2 = str_now
    time_1 = datetime.strptime(time_1, display) # 轉成<class 'datetime.datetime'>
    time_2 = datetime.strptime(time_2, display) # 轉成<class 'datetime.datetime'>
    time_interval = time_2 - time_1
    print(time_interval.days)
    print(time_interval.seconds)
    days = time_interval.days
    seconds = time_interval.seconds%60
    minutes = int(time_interval.seconds/60)%60
    hours = int(int(time_interval.seconds/60)/60)
    msg = ''
    if days!=0:
        msg = msg+str(days)+'天'
    if hours!=0:
        msg = msg+str(hours)+'小時'
    if minutes!=0:
        msg = msg+str(minutes)+'分'  
    if seconds!=0:
        msg = msg+str(seconds)+'秒'  
    return msg


# 算兩個時間的時間差
def CountTimeDiff(time_1="", time_2="", display="%Y-%m-%d"):
    # 抓現下時間
    today = date.today()
    str_today = today.strftime(display) 
    if time_1=="" :
        time_1 = str_today
    if time_2=="" :
        time_2 = str_today
    time_1 = datetime.strptime(time_1, display) # 轉成<class 'datetime.datetime'>
    time_2 = datetime.strptime(time_2, display) # 轉成<class 'datetime.datetime'>
    time_interval = time_2 - time_1
    # 回傳的Type會是 <class 'datetime.delta'>
    # 要看相差幾天 ---> time_interval.days
    return time_interval

# =====================================================================================================================  Divide List
# 把大List確實依照每num個分割成一個小List，會有分配非常不均的可能
def DivideList(BigList, num):
    SubLists = [ BigList[i:i+num] for i in range(0, len(BigList), num) ]
    return SubLists

# 把大List用Pooling的分方式，分割成各個個小List
def DivideList(BigList, num): # 每個小List的元素量雖然會參考num，但仍會自行調整，不會有分配非常不均的可能
    threadsNum = int(len(BigList)/num)+1 if int(len(BigList)%num)>0 else int(len(BigList)/num)

    SubLists = []
    for i in range(threadsNum):
        SubLists.append([]) 

    for j in range(num):
        for i in range(len(SubLists)):
            k = i + (j*threadsNum)
            if k <= (len(BigList)-1):
                SubLists[i].append( BigList[k] ) 
    return SubLists

# ====================================================================================================================== About Txt File
# 讀只有單一行的txt檔內容
def ReadTxtFileSingleRow(Path='pwd', FileName='NextBeginFrom.txt'):
    if Path=='pwd':
        # 取得當目錄
        r = os.popen("pwd") # 執行cmd
        text = r.read()
        r.close()
        Path = text[:(len(text)-1)] # -1 是因為它的結果最後會多一個換行符號
    # 讀取txt檔
    path = Path + '/' + str(FileName)
    f = open(path, 'r')
    Content = f.read()
    Content = Content[:len(Content)-1]
    f.close()
    print("===")
    print(Content)
    print("===")
    return Content
    
# 覆蓋寫入Txt檔
def WriteTxtFile(Path='pwd', FileName='NextBeginFrom.txt', msg='測試'):
    if Path=='pwd':
        # 取得當目錄
        r = os.popen("pwd") # 執行cmd
        text = r.read()
        r.close()
        Path = text[:(len(text)-1)] # -1 是因為它的結果最後會多一個換行符號
    # 讀取txt檔
    path = Path + '/' + str(FileName)
    with open(path,"w") as f:
        f.write(msg)  # 自带文件关闭功能，不需要再写f.close()
    print("目標："+str(path))
    print("寫入txt的文字訊息為\n")
    print(msg)


# ======================================================================================================== Pandas相關 常用功能與筆記
# ########################################## 取得有哪些班級   
# StuClasses = list(set(dfStuClasses['Class']))
# StuClasses.sort(reverse=False)
# buffer = []
# for i in range(len(StuClasses)):
#     buffer.append(str(StuClasses[i]))
# StuClasses = buffer    
    
# ########################################## 宣告一個Dataframe, 定義好欄位 (columns=[]也可以) 
# unregistered_CardID  = pd.DataFrame(columns = ['Timming(Date)', 'CardID', 'Choose', 'WaterVolume', 'WaterTemp', 'Dispenser'])

# ########################################## 往Dataframe新增一筆資料
# AppendData = {
#     'Timming(Date)': Find_unregistered.iloc[i]['Timming(Date)'], 
#     'CardID': Find_unregistered.iloc[i]['CardID'], 
#     'Choose': Find_unregistered.iloc[i]['Choose'], 
#     'WaterVolume': Find_unregistered.iloc[i]['WaterVolume'], 
#     'WaterTemp': Find_unregistered.iloc[i]['WaterTemp'], 
#     'Dispenser': Find_unregistered.iloc[i]['Dispenser']
# }
# unregistered_CardID = unregistered_CardID.append(AppendData, ignore_index=True, sort=False)    

# ########################################## 調換dataframe的欄位順序
# Faculty_raw_data = Faculty_raw_data[['Timming(Date)', 'Choose', 'WaterTemp','WaterVolume', 'Dispenser', 'CardID', 'Identity']]    

# ########################################## 把dataframe的index重設
# Faculty_raw_data  = Faculty_raw_data.reset_index(drop=True)
    
# ########################################## 新增一欄'Date', 依據是把'Timming'只取日期的部分
# dfReadData['Date']= dfReadData['Timming'].apply(lambda x: x[:10])

# ########################################## 把'Timming'去除掉時間的部分
# dfReadData['Time']= dfReadData['Timming'].apply(lambda x: x[11:])

# ########################################## 給dataframe某一欄以較複雜的條件進行校正
# 定義好校正的方程式
# def VolumneCorrection(WaterVolume): 
#     if WaterVolume<0 :
#         WaterVolume = float(int( WaterVolume + 1000000 ))
#     if WaterVolume>10000 :
#         WaterVolume = float(int( 1000000-WaterVolume ))    
#     return WaterVolume
# 把pandas的警告關掉
# pd.options.mode.chained_assignment = None
# 校正WaterVolume
# dfReadData['WaterVolume']= dfReadData['WaterVolume'].apply(VolumneCorrection) 

# ########################################## 將dataframe符合條件的資料濾出來
# Filter21 = effectiveExcel.loc[(FinalData['Timming(Date)']==Date)&(FinalData['SchoolID']!='')]

# ########################################## 把整欄資料做型態轉換
# df['學號'] = df['學號'].astype(str)

# ########################################## 針對某一column做正則表達式，將dataframe符合條件的資料濾出來
# df = df[df['學號'].str.contains("^107")] # 開頭
# df = df[df['學號'].str.contains("023$")] # 結尾

# ########################################## 濾出該欄位不為空的資料
# Filter11 = Filter11[Filter11.SchoolID.notnull()]

# ########################################## 比對兩個dataframe做合併
# dfReadData=dfReadData.join(All_DistinguishList.set_index('Input'), on='CardID',lsuffix='_l', rsuffix='_r')

# ########################################## 把對不上的資料(NaN), 以"unregistered"表示
# fillValues = {'Class':"unregistered", 'Gender':"unregistered", 'Grade':"unregistered", 
#               'Identity':"unregistered", 'Room':"unregistered", 'SchoolID':"unregistered"}
# dfReadData = dfReadData.fillna(value=fillValues)

# ########################################## 把dataframe輸出成xlsx檔
# FileName = 'Failed_raw_data_'+timeString+'.xlsx'
# writer = pd.ExcelWriter(FileName, engine='xlsxwriter')
# Failed_raw_data.to_excel(writer, sheet_name='Failed_raw_data', index=False)
# writer.save()

