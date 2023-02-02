import sys
import pymongo
from pymongo import MongoClient
import pandas as pd
import csv
import os
import time

mongo_url_01 = "mongodb://admin:bmwee8097218@140.118.122.115:30415/"
mongo_url_02 = "mongodb://admin:bmwee8097218@140.118.122.115:30415/"
mongo_url_03 = "mongodb://admin:bmwee8097218@140.118.122.115:30708/"

def Membersdata(DB, Collection, Search={}):
    #Membersdata(DB="xinxing_dispenser",Collection="members_data_test",Search={'Class':'701'})
    
    global mongo_url_01,mongo_url_02
    try:
        conn = MongoClient(mongo_url_01) 
        db = conn[DB]
        collection = db[Collection]
        cursor=collection.find(Search)
        data=[d for d in cursor]
    except:
        conn = MongoClient(mongo_url_02) 
        db = conn[DB]
        collection = db[Collection]
        cursor=collection.find(Search)
        data=[d for d in cursor]
    if data==[]:
        return False
    else:
        return data


def Membersdata_sequence(DB, Collection, Class):
    data_2=Membersdata(DB=DB,Collection=Collection,Search={'Class':str(Class),'Identity':'Student'})
    temp=0
    SchoolID=[]
    data=[]
    while len(data_2)>0:
        num=999
        for i in range(len(data_2)):
            if num>int(data_2[i]["Number"]):
                num=int(data_2[i]["Number"])
                p=i
        SchoolID.append(str(data_2[p]["SchoolID"]))
        data.append(data_2[p])
        data_2.remove(data_2[p])
    additional_card=Membersdata(DB, "members_addtional_card_test",Search={})
    for i in range(len(additional_card)):
        try:
            path=SchoolID.index(str(additional_card[i]['SchoolID']))
            error=0
        except:
            error=1
        if error==0 and additional_card[i]['primaryCard']!=additional_card[i]['additionalCard']:
            insertdata={
                            'SchoolID':data[path]['SchoolID'],
                            'Identity':data[path]['Identity'],
                            'Name':data[path]['Name'],
                            'Gender':data[path]['Gender'],
                            'CardID':additional_card[i]['additionalCard'],
                            'Class':data[path]['Class'],
                            'Weight':data[path]['Weight'],
                            'DailyShouldDrink':data[path]['DailyShouldDrink'],
                            'Height':data[path]['Height'],
                            'Number':data[path]['Number']
                        }
            data.insert( 1+path, insertdata)
            SchoolID.insert(1+path,str(data[path]['SchoolID']))
    return data
def teadatamerge(data):
    alldata=[i for i in range(len(data))]
    for i in range (len(data)):
        data=data.where(data.notnull(), None)
        try:
            if len(str(int(data.at[i,'校務行政'])))==3:
                    SchoolID="00"+str(int(data.at[i,"校務行政"]))
                    CardID="00"+str(int(data.at[i,"校務行政"]))
            elif len(str(int(data.at[i,'校務行政'])))==2:
                    SchoolID="000"+str(int(data.at[i,"校務行政"]))
                    CardID="000"+str(int(data.at[i,"校務行政"]))
            else:
                SchoolID=str(int(data.at[i,"校務行政"]))
                CardID=str(int(data.at[i,"校務行政"]))
            alldata[i]={
                             'SchoolID':str(SchoolID),
                             'CardID':str(CardID),
                             'Class':str("unregistered"),
                             'DailyShouldDrink':float(0.0),
                             'Identity':"Faculty",
                             'Name':str(data.at[i,"姓名"]),
                             'Gender':str("unregistered"),
                             'Height':float(0.0),
                             'Weight':float(0.0),
                             }
        except:
            alldata="Error"
    return alldata
def studatamerge(data,data_1):

    for i in range(len(data)):
        data=data.where(data.notnull(), None)
    for i in range(len(data_1)):
        data_1=data_1.where(data_1.notnull(), None)
    print(data)
    try:
        #df=data.merge(data_1.set_index('學號'), on=('學號'))
        if len(data)>len(data_1):
            df=data.join(data_1.set_index('學號'), on='學號',lsuffix='_l', rsuffix='_r')
        else:
            df=data_1.join(data.set_index('學號'), on='學號',lsuffix='_l', rsuffix='_r')
        df=df.where(df.notnull(), "None")
        err=0
    except:
        err=1
    if err==0:
        alldata=[i for i in range(len(df))]
        for i in range(len(df)):
            SchoolID=df.at[i,'學號']
            try:
                CardID=df.at[i,'卡片內碼_l']
            except:
                try:
                    CardID=df.at[i,'卡片內碼_r']
                except:
                    try:
                        CardID=df.at[i,'卡片內碼']
                    except:
                        CardID="None"
               
            try:
                Name=df.at[i,'姓名_l']
            except:
                try:
                    Name=df.at[i,'姓名_r']
                except:
                    try:
                        Name=df.at[i,'姓名']
                    except:
                        Name="None"
            try:
                Class=df.at[i,'班級_l']
            except:
                try:
                    Class=df.at[i,'班級_r']
                except:
                    try:
                        Class=df.at[i,'班級']
                    except:
                        Class="None"
            try:
                Height=df.at[i,'身高_l']
            except:
                try:
                    Height=df.at[i,'身高_r']
                except:
                    try:
                        Height=df.at[i,'身高']
                    except:
                        Height="None"
            try:
                Weight=df.at[i,'體重_l']
            except:
                try:
                    Weight=df.at[i,'體重_r']
                except:
                    try:
                        Weight=df.at[i,'體重']
                    except:
                        Weight="None"
            try:
                Number=df.at[i,'座號_l']
            except:
                try:
                    Number=df.at[i,'座號_r']
                except:
                    try:
                        Number=df.at[i,'座號']
                    except:
                        Number="None"
            try:
                Gender=df.at[i,'性別_l']
            except:
                try:
                    Gender=df.at[i,'性別_r']
                except:
                    try: 
                        Gender=df.at[i,'性別']
                    except:
                        Gender="None"
            if CardID!="None" :
                if len(str(int(CardID)))==9:
                    CardID="0"+str(int(CardID))
                else:
                    CardID=str(int(CardID))
            else:
                CardID="unregistered"
            if Name!="None":
                Name=str(Name)
            else:
                Name="unregistered"
            if Class!="None":
                Class=str(Class)
            else:
                Class="unregistered"
            if Height!="None":
                Height=float(Height)
            else:
                Height=0.0
            if Weight!="None":
                Weight=float(Weight)
            else:
                Weight=0.0
            if Number!="None":
                if len(str(Number))==1:
                    Number="0"+str(Number)
                else:
                    Number=str(Number)
            else:
                Number="unregistered"
            if Gender!="None":
                if Gender=="女"or str(Gender)=="2":
                    Gender="Female"
                else:
                    Gender="Male"
            else:
                Gender="unregistered"
            DailyShouldDrink=Weight*40
            alldata[i]={
                     'SchoolID':SchoolID,
                     'CardID':CardID,
                     'Class':Class,
                     'DailyShouldDrink':DailyShouldDrink,
                     'Identity':"Student",
                     'Name':Name,
                     'Gender':Gender,
                     'Height':Height,
                     'Weight':Weight,
                     "Number":Number}
           
    else:
        alldata="Error"

    return alldata
def updateInDB(DB,Collection,selector, data):
    
    try:
        conn = MongoClient(mongo_url_01) 
        db = conn[DB]
        collection = db[Collection]
        collection.update_one(selector, data)
       # print("31383: update success")
    except:
        try:
            conn = MongoClient(mongo_url_02)
            db = conn[DB]
            collection = db[Collection]
            collection.update_one(selector, data)
           # print("30415: update success")
        except:
            try:
                conn = MongoClient(mongo_url_03)
                db = conn[DB]
                collection = db[Collection]
                collection.update_one(selector, data)
           #     print("30708: update success")
            except:
                print("no one success to write data into DB !")
def ReadFromDB(DB,Collection,SchoolID):
    global mongo_url_01
    conn = MongoClient(mongo_url_01) 
    db = conn[DB]
    collection = db[Collection]
    cursor = collection.find({"SchoolID":SchoolID}) #此處須注意，其回傳的並不是資料本身，你必須在迴圈中逐一讀出來的過程中，它才真的會去資料庫把資料撈出來給你。
    data = [d for d in cursor] #這樣才能真正從資料庫把資料庫撈到python的暫存記憶體中。
    if data==[]:
        return False
    else:
        return True
    
    new_data={
        "client_ip_address":"123",
        ""
        
        
        
        
        }
def WriteInDB(DB,Collection,new_data):
    try:
        conn = MongoClient(mongo_url_01) 
        db = conn[DB]
        collection = db[Collection]
        collection.insert(new_data)
    except:
        try:
            conn = MongoClient(mongo_url_02)
            db = conn[DB]
            collection = db[Collection]
            collection.insert(new_data)
        except:
            try:
                conn = MongoClient(mongo_url_03)
                db = conn[DB]
                collection = db[Collection]
                collection.insert(new_data)
            except:
                print("no one success to write data into DB !")
def dataupdate(DB,Collection,alldata):
    for i in range(len(alldata)):
        SchoolID=str(alldata[i]['SchoolID'])
        Name=str(alldata[i]['Name'])
        Class=str(alldata[i]['Class'])
        Height=float(alldata[i]['Height'])
        Weight=float(alldata[i]['Weight'])
        DailyShouldDrink=float(alldata[i]['DailyShouldDrink'])
        Number=str(alldata[i]['Number'])
        CardID=str(alldata[i]['CardID'])
        Gender=str(alldata[i]['Gender'])
        TodayDrinkRefreshDate = time.strftime("%Y-%m-%d", time.localtime()) # 轉成字串
       # result = ReadFromDB(DB,Collection,SchoolID)
        result=Membersdata(DB,Collection,{"SchoolID":SchoolID,"Identity":"Student"})
        if result!=False:
            if Class=="unregistered" or Height==0.0 or Weight==0.0 or Number=="unregistered" or CardID=="unregistered" :
                 data=Membersdata(DB,Collection,Search={'SchoolID':SchoolID})
                 if CardID=="unregistered":
                     CardID=data[0]["CardID"]
                 if Class=="unregistered":
                     Class=data[0]["Class"]
                 if Height==0.0:
                     Height=float(data[0]["Height"])
                 if Weight==0.0:
                     Weight=float(data[0]["Weight"])
                     DailyShouldDrink=Weight*40
                 if Number=="unregistered":
                     Number=data[0]["Number"]
            selector = { "SchoolID": SchoolID }       
            data = { "$set": { "Name":Name, "Class":Class, "Height":Height, "Weight":Weight, "DailyShouldDrink":DailyShouldDrink,"Number":Number,"CardID":CardID,"TodayDrinkRefreshDate":TodayDrinkRefreshDate } }
            updateInDB(DB,Collection,selector, data)
        else:
            new_data = {"SchoolID": SchoolID, "Identity":"Student", "Name":Name, "EnrollYear":"none", "Gender":Gender, "CardID":CardID, "Class":Class, "Weight":Weight, "DailyShouldDrink":DailyShouldDrink, "TodayDrinkRefreshDate":TodayDrinkRefreshDate, "TodayDrink":0.0, "Height":Height,"Number":Number}
            WriteInDB(DB,Collection,new_data)
def dataupdata_grade(DB,Collection,alldata):
    conn = MongoClient(mongo_url_01) 
    db = conn[DB]
    collection = db[Collection]
    cursor = collection.find({"Identity":"Student"}) #此處須注意，其回傳的並不是資料本身，你必須在迴圈中逐一讀出來的過程中，它才真的會去資料庫把資料撈出來給你。
    data = [d for d in cursor]
    schoolid=[i for i in range(len(alldata))]
    for i in range (len(alldata)):
        schoolid[i]=str(alldata[i]['SchoolID'])
    for i in range(len(data)):
        SchoolID = str(data[i]['SchoolID'])
        if (SchoolID in schoolid)==False:
            selector = { "SchoolID": str(SchoolID) }
            data2 = { "$set": { "Identity":"Alumni" } }
            updateInDB(DB,Collection,selector, data2)
def teadataupdate(DB,Collection,alldata):
    for i in range(len(alldata)):
        SchoolID=str(alldata[i]['SchoolID'])
        Name=str(alldata[i]['Name'])
        CardID=str(alldata[i]['CardID'])
        TodayDrinkRefreshDate = time.strftime("%Y-%m-%d", time.localtime()) # 轉成字串
        result=Membersdata(DB,Collection,{"SchoolID":SchoolID,"Identity":"Faculty"})
      #  result = ReadFromDB(DB,Collection,SchoolID)

        if result!=False:
            selector = { "SchoolID": SchoolID }
            data = { "$set": { "Name":Name,"CardID":CardID,"TodayDrinkRefreshDate":TodayDrinkRefreshDate } }
            updateInDB(DB,Collection,selector, data)
        else:
            new_data = {"SchoolID": SchoolID, "Identity":"Faculty", "Name":Name, "EnrollYear":"none", "Gender":"unregistered", "CardID":CardID, "Class":"unregistered", "Weight":0.0, "DailyShouldDrink":0.0, "TodayDrinkRefreshDate":TodayDrinkRefreshDate, "TodayDrink":0.0, "Height":0.0}
            WriteInDB(DB,Collection,new_data)
def datadoublepath(data):
    dataSchoolID=[i for i in range(len(data))]
    for i in range(len(data)):
        dataSchoolID[i]=data[i]['SchoolID']
    newSchoolID=[]
    additional_card=Membersdata("xinxing_dispenser", "members_addtional_card_test",Search={})
    SchoolID=[i for i in range(len(additional_card))]
    for i in range(len(additional_card)):
        SchoolID[i]=additional_card[i]['SchoolID']
        if not SchoolID[i] in newSchoolID :
            if SchoolID[i] in dataSchoolID:
                newSchoolID.append(str(SchoolID[i]))
    for i in range(len(newSchoolID)):
            newSchoolID.append(int(SchoolID.count(newSchoolID[i])))
            try:
                path=dataSchoolID.index(str(newSchoolID[i]))
                error=0
            except:
                error=1
            if error==0:
                newSchoolID.append(path)
    return(newSchoolID)

###相對檔案路逕###
#path=os.getcwd()
#path=str(os.path.dirname(path))
#path=path+"\\xinxing_pyqt_config_doc\\xinxingQt_config.csv"
#################



			
		
	