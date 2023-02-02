from PyQt5 import QtWidgets, QtGui, QtCore
import UI
from UI import Ui_MainWindow
import sys
from bson.objectid import ObjectId #這東西在透過ObjectID去尋找的時候會用到
# 這邊專門給我的 Module使用

from importlib import reload
# # 將project_home指定為你的專案路徑 要在同一個Disk裡面
# project_home = u'/Keep_the_desktop_clean/save_jupyter_python_code/ForXinxingProject/MyModule'
# if project_home not in sys.path:
#     sys.path = [project_home] + sys.path
# # 可以讀取整個工具組，也可以讀取特定function
sys.path.append(r'xinxing_functions')
import RJ
RJ = reload(RJ) # 因為我的module也一直在改，所以就...每次執行都要重新import
# print(RJ.Owner)

import XinxingPG
XinxingPG = reload(XinxingPG) # 因為我的module也一直在改，所以就...每次執行都要重新import
print(XinxingPG.Owner)

import datetime
from datetime import datetime
import time
import tkinter 
from tkinter import ttk
from tkinter import messagebox
import tkinter.font as tkFont
import string
import random
import threading
import XinxingQt as xqt
import subfunction as sub
####################
from PyQt5.QtWidgets import *
from PyQt5 import QtCore, QtGui
from PyQt5.QtGui import *
from PyQt5.QtCore import *
import pandas as pd
import requests

global DB
DB="xinxing_dispenser"
global Colleciton
Collection="members_data"
global CollectionMAC
CollectionMAC="members_addtional_card_test"
class SecondWindow(QDialog,QWidget,UI.Ui_SecondWindow):
    global second_check
    second_check=""
    def __init__(self):
        super(SecondWindow,self).__init__()
        self.setupUi(self)
        self.pushButton.clicked.connect(self.add_update)
    def Check(self,message,message_2):
        data=QMessageBox.information(self, message, message_2,QMessageBox.Yes|QMessageBox.No , QMessageBox.No)
        #16384=Yes
        #65536=No
        return data
    def Error(self,message="error"):
        QMessageBox.critical(self,'嚴重警告',message, QMessageBox.Close  , QMessageBox.Close )
    def add_update(self):
        global second_check
        SchoolID=self.lineEdit.text()
        Name=self.lineEdit_2.text()
        CardID=self.lineEdit_3.text()
        Class=self.lineEdit_4.text()
        Weight=self.lineEdit_5.text()
        Height=self.lineEdit_6.text()
        Number=self.lineEdit_7.text()
        Gender=self.lineEdit_8.text()
        Errorcode=""
        print(second_check)
        if len(SchoolID)!=8:
            Errorcode=Errorcode+("學號格式有誤\n")
        elif second_check!="edit":
            print(second_check)
            second_check=""
            data=xqt.Membersdata(DB,Collection,Search={'SchoolID':SchoolID,'Identity':'Student'})
            if data!=False:
                SchoolID="Error"
                Errorcode=Errorcode+("學號重複\n")
        elif second_check=="edit":
             second_check=""
        if len(CardID)==10 :
            data=xqt.Membersdata(DB,Collection,Search={'CardID':CardID,'Identity':'Student'})
            if data!=False :
                if data[0]['SchoolID']!=SchoolID:
                    CardID="Error"
                    Errorcode=Errorcode+("卡號重複\n")
        elif CardID=="unregistered":
            CardID=CardID
        else:
            CardID="Error"
            Errorcode=Errorcode+("卡號有誤\n")
        if Name=="":
            Name="unregistered"
        else:
          result=0
          for z in range (len(Name)):
             if  Name[z]>= u'\u4e00' and Name[z]<=u'\u9fff':
                 result=result
             else:
                 result=result+1
          if result>0:
                Name="Error"
                Errorcode=Errorcode+("姓名非中文\n")
          else:
                Name=Name         
        if Class=="":
            Class="unregistered"
        if Weight=="": 
            Weight=0.0
        if Height=="": 
            Height=0.0
        if Number=="":
            Number="0"
        elif len(Number)==1:
            Number="0"+str(Number)
        if Gender=="":
            Gender="unregistered"
        elif Gender=="男" or Gender=="2" or Gender=="Male":
            Gender="Male"
        elif Gender=="女" or Gender=="1" or Gender=="Female":
            Gender="Female"
        else:
            Gender="Error"
            Errorcode=Errorcode+("性別非男、女\n")
        DailyShouldDrink=float(Weight)*40
        alldata=[i for i in range(1)]
        alldata[0]={
            'SchoolID':str(SchoolID),
            'CardID':str(CardID),
            'Class':str(Class),
            'DailyShouldDrink':float(DailyShouldDrink),
            'Identity':"Student",
            'Name':str(Name),
            'Gender':str(Gender),
            'Height':float(Height),
            'Weight':float(Weight),
            "Number":str(Number)
            }
       
        if Errorcode!="":
                self.Error(Errorcode)
        else:
            checkcode="學號:"+SchoolID+"\n姓名:"+Name+"\n卡號:"+CardID+"\n班級:"+Class+"\n性別:"+Gender+"\n身高:"+str(Height)+"\n體重:"+str(Weight)+"\n座號:"+Number
            check=self.Check("確認上傳",checkcode)
            if check==16384:
                xqt.dataupdate(DB,Collection,alldata)
                self.close()
        
    def setup_add(self):
        global second_check
        self.pushButton.setText("上傳新增")
        second_check=""
    def setup_edit(self,data):
        global second_check
        self.lineEdit.setText(data['lineEdit'])
        self.lineEdit.setFocusPolicy(QtCore.Qt.NoFocus)
        self.lineEdit_2.setText(data['lineEdit_2'])
        self.lineEdit_3.setText(data['lineEdit_3'])
        self.lineEdit_3.setFocusPolicy(QtCore.Qt.NoFocus)
        self.lineEdit_4.setText(data['lineEdit_4'])
        self.lineEdit_5.setText(data['lineEdit_5'])
        self.lineEdit_6.setText(data['lineEdit_6'])
        self.lineEdit_7.setText(data['lineEdit_7'])
        self.lineEdit_8.setText(data['lineEdit_8'])
        self.lineEdit_8.setFocusPolicy(QtCore.Qt.NoFocus)
        second_check=data['Check']
        self.pushButton.setText("檢驗格式並更新")
class ThirdWindow(QDialog,QWidget,UI.Ui_ThirdWindow):
    dialogSignel=pyqtSignal(int)
    def __init__(self):
        super(ThirdWindow,self).__init__()
        self.setupUi(self)
        self.pushButton.clicked.connect(self.add_update)
    def Check(self,message,message_2):
        data=QMessageBox.information(self, message, message_2,QMessageBox.Yes|QMessageBox.No , QMessageBox.No)
        #16384=Yes
        #65536=No
        return data
    def Error(self,message="error"):
        QMessageBox.critical(self,'嚴重警告',message, QMessageBox.Close  , QMessageBox.Close )
    def add_update(self):
        SchoolID=str(self.lineEdit.text())
        Name=self.lineEdit_2.text()
        CardID=self.lineEdit_3.text()
        print(CardID)
        Errorcode=""
        if len(SchoolID)==5:
            SchoolID=SchoolID
            print(SchoolID)
        elif len(SchoolID)==2:
            SchoolID="000"+SchoolID
        elif len(SchoolID)==3:
            SchoolID="00"+SchoolID
        else:
            SchoolID="Error"
            Errorcode=Errorcode+("教職員代碼有誤\n")
        if Name=="":
            Name="unregistered"
        else:
          result=0
          for z in range (len(Name)):
             if  Name[z]>= u'\u4e00' and Name[z]<=u'\u9fff':
                 result=result
             else:
                 result=result+1
          if result>0:
                Name="Error"
                Errorcode=Errorcode+("姓名非中文\n")
          else:
                Name=Name
        if CardID=="":
            CardID=SchoolID
        else:
            if len(CardID)==5 or len(CardID)==10:
                data=xqt.Membersdata(DB,Collection,Search={'CardID':CardID})
                if data!=False :
                    if data[0]['SchoolID']!=SchoolID:
                        CardID="Error"
                        Errorcode=Errorcode+("卡號重複\n")
                    else:
                        CardID=CardID
                else:
                    CardID=CardID
            else:
                CardID="Error"
                Errorcode=Errorcode+("卡號有誤\n")
        alldata=[i for i in range(1)]
        alldata[0]={
            'SchoolID':str(SchoolID),
            'CardID':str(CardID),
            'Name':str(Name)
            }
        if Errorcode!="":
                self.Error(Errorcode)
        else:
            checkcode="教職員代碼:"+SchoolID+"\n姓名:"+Name+"\n卡號:"+CardID
            check=self.Check("確認更新",checkcode)
            if check==16384:
                xqt.teadataupdate(DB,Collection,alldata)
                self.close() 
    def setup_add(self):
        self.pushButton.setText("上傳新增")
    def setup_edit(self,data):
        self.lineEdit.setText(data['lineEdit'])
        self.lineEdit.setFocusPolicy(QtCore.Qt.NoFocus)
        self.lineEdit_2.setText(data['lineEdit_2'])
        self.lineEdit_3.setText(data['lineEdit_3'])
        self.lineEdit_3.setFocusPolicy(QtCore.Qt.NoFocus)
        self.pushButton.setText("檢驗格式並更新")
class MainWindow_controller(QtWidgets.QMainWindow):
    dialogSignel=pyqtSignal(int)
    def __init__(self):
        super().__init__() # in python3, super(Class, self).xxx = super().xxx
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.setui()
        self.ui.action.triggered.connect(self.example)
        #########水量報表#########
        self.ui.calculate.clicked.connect(self.Calculate0)
        self.ui.startbutton.clicked.connect(self.starbutton)
        self.ui.endbutton.clicked.connect(self.endbutton)
        self.ui.excelbutton.clicked.connect(self.searchexcelpath)
        #########學生資料#########
        self.ui.stucom.activated.connect(self.stucom)
        self.ui.stucom_2.activated.connect(self.stucom)
        self.ui.stusearch.clicked.connect(self.stusearch)
        self.ui.studelet.clicked.connect(self.studelet)
        self.ui.stusearview.clicked.connect(self.stusearview)
        self.ui.stusearview_2.clicked.connect(self.stusearview_2)
        self.ui.stuupdate.clicked.connect(self.stuupdata)
        self.ui.stuupdate_2.clicked.connect(self.stuupdata_grade)
        self.ui.stuedit.clicked.connect(self.stuedit)
        self.ui.sturadd.clicked.connect(self.stuadd)
        self.ui.subfuction.clicked.connect(self.stusubfuction)
        #########教職員資料########
        #self.ui.teaupdate.clicked.connect(self.teaupdate)
        self.ui.teasearch.clicked.connect(self.teabutton)
        self.ui.teadelet.clicked.connect(self.teadelet)
        #self.ui.teasearview.clicked.connect(self.teasearview)
        self.ui.teaedit.clicked.connect(self.teaedit)
        self.ui.teaadd.clicked.connect(self.teaadd)
        self.ui.teasubfunction.clicked.connect(self.teasubfunction)
    def setui(self):
        self.ui.startdate.setText(RJ.AddOrMinusDays(RJ.Get_date_of_today(display="%Y-%m-%d"), delta=-7, display="%Y-%m-%d"))
        self.ui.startdatewidge.setMinimumDate(datetime.strptime("2021-10-25", "%Y-%m-%d"))
        self.ui.startdatewidge.setMaximumDate(datetime.strptime(RJ.Get_date_of_yesterday(display="%Y-%m-%d"), "%Y-%m-%d"))
        self.ui.enddatewidge.setMinimumDate(datetime.strptime("2021-10-25", "%Y-%m-%d"))
        self.ui.enddatewidge.setMaximumDate(datetime.strptime(RJ.Get_date_of_yesterday(display="%Y-%m-%d"), "%Y-%m-%d"))
        self.ui.enddate.setText(RJ.Get_date_of_yesterday(display="%Y-%m-%d"))
        self.ui.excelpath.setText("F:/")
        self.ui.dataedit.setText("7")
        self.ui.startdatewidge.setVisible(False)
        self.ui.enddatewidge.setVisible(False)
        self.stucom()
        self.teabutton()
    def Check(self,message,message_2):
        data=QMessageBox.information(self, message, message_2,QMessageBox.Yes|QMessageBox.No , QMessageBox.No)
        #16384=Yes
        #65536=No
        return data
    def Error(self,message="error"):
        QMessageBox.critical(self,'嚴重警告',message, QMessageBox.Close  , QMessageBox.Close )
    def example(self):
       SavePath,_ = QFileDialog.getSaveFileName(None,  "文件保存","C:\\","xlsx (*.xlsx *.xls )")
       data = pd.DataFrame(columns = ['學號', '姓名', '班級', '座號', '性別', '身高', '體重', '卡片內碼']) 
       
       try:
           #SavePath=SavePath.replace( '/', "\\") 
           result=pd.ExcelWriter(SavePath)
           data.to_excel(result, sheet_name='學生資料',index=False)
           result.save()
       except:
           pass
#########教職員資料#############
    def openthird(self,data):
        self.dialog=ThirdWindow()
        self.dialog.show()
        if data=="":
            self.dialog.setup_add()
        else:
            self.dialog.setup_edit(data)
    def teaadd(self):
        self.openthird("")
    def teaedit(self):
        tabledata=self.ui.teatab.currentIndex().row()
        model = self.ui.teatab.model()
        if tabledata!=-1:
            SchoolID=model.data(model.index(tabledata,0))
            Name=model.data(model.index(tabledata,2))
            CardID=model.data(model.index(tabledata,4))
            data={
                'lineEdit':SchoolID,
                'lineEdit_2':Name,
                'lineEdit_3':CardID
                }
            self.openthird(data)
        else:
            self.Error("請先選擇要編輯的資料")        
    def teasearview(self):
        #教師批量讀
        self.ui.teaupedit.setText("")
        filename=QFileDialog.getOpenFileName(None,'Open file',"C:","xlsx (*.xlsx *.xls )")
        text=filename[0]
        self.ui.teaupedit.setText(text)
        if  len(filename)>=0:
            try :
                data=pd.read_excel(str(filename[0]))
                df=pd.DataFrame(data)
            except:
                self.Error("檔案格式有誤")
        alldata=xqt.teadatamerge(df)
        if alldata!="Error" and len(alldata)!=0:
            self.ui.tearowreload(alldata)
        else:
            self.Error("檔案內容有誤")
    def teadelet(self):
        tabledata=self.ui.teatab.currentIndex().row()
        model = self.ui.teatab.model()
        if tabledata!=-1:
            message="是否要刪除\n"+"教職員代碼:"+str(model.data(model.index(tabledata,0)))+"\n姓名:"+str(model.data(model.index(tabledata,2)))+"\n卡號:"+str(model.data(model.index(tabledata,4)))
            check=self.Check("確認刪除",message)  
            if check==16384:
                SchoolID=model.data(model.index(tabledata,0))
                CardID=model.data(model.index(tabledata,4))
                data_2=xqt.Membersdata(DB, Collection,Search={'SchoolID':SchoolID,'Identity':'Faculty','CardID':CardID})
                if data_2!=False:
                    selector = { "SchoolID": str(model.data(model.index(tabledata,0))),"Identity":"Faculty" }
                    data2 = { "$set": { "Identity":"Alumni" } }
                    xqt.updateInDB(DB,Collection,selector, data2)
                try:
                    data=xqt.Membersdata(DB, CollectionMAC,Search={'SchoolID':SchoolID,'additionalCard':CardID})
                    error=0
                except:
                    error=1
                if error==0 and data!=False:
                    num=int(len(xqt.Membersdata(DB, CollectionMAC,Search={'SchoolID':SchoolID})))
                    if num==2:
                        Condition={"SchoolID":model.data(model.index(tabledata,0))}
                        RJ.SC_DeleteDataFromDB(DB,CollectionMAC,Condition,"Many")
                    elif num!=2 and data_2!=False:
                        Condition={"SchoolID":model.data(model.index(tabledata,0))}
                        RJ.SC_DeleteDataFromDB(DB,CollectionMAC,Condition,"Many")
                    else:
                        Condition={"SchoolID":model.data(model.index(tabledata,0)),"additionalCard":model.data(model.index(tabledata,4))}
                        RJ.SC_DeleteDataFromDB(DB,CollectionMAC,Condition,"One")
                self.teabutton()
        else:
            self.Error("請先選擇要刪除的資料")
       
    def teabutton(self):
        SchoolID=str(self.ui.teasearedit.text())
        if self.ui.teasearedit.text()!="":
            Error=0
            data=xqt.Membersdata(DB, Collection,Search={'Identity':'Faculty','SchoolID':SchoolID})
            if data==False:
                Error=0
                data=xqt.Membersdata(DB, Collection,Search={'Identity':'Faculty','SchoolID':"00"+SchoolID})
                if data==False:
                    Error=1
                    self.ui.teasearedit.setText("")
                    self.Error("查無代碼")
        else:
            Error=0
            data=xqt.Membersdata(DB, Collection,Search={'Identity':'Faculty'})
        if Error==0:
            additional_card=xqt.Membersdata(DB, CollectionMAC,Search={})
            SchoolID=[i for i in range(len(data))]
            for i in range(len(data)):
                SchoolID[i]=str(data[i]['SchoolID'])
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
                                    'Height':data[path]['Height']
                                }
                    data.insert( 1+path, insertdata)
                    SchoolID=[i for i in range(len(data))]
                    for i in range(len(data)):
                        SchoolID[i]=str(data[i]['SchoolID'])
                   
            self.ui.teasearedit.setText("")
            self.ui.tearowreload(data)
            self.teasetspan(data)
    def teaupdate(self):
        #教師批量上傳
        if self.ui.teaupedit.text()!="":
            Errorcode=""
            row = self.ui.teatab.rowCount()
            alldata=[i for i in range(row)]
            for i in range(int(row)):
                for j in range(5):
                    item = self.ui.teatab.item(i, j)
                    data = str(item.text())
                    if j==0:
                        if len(data)==5 or data=="unregistered":
                            SchoolID=data
                        else:
                            SchoolID="Error"
                            Errorcode=Errorcode+("第"+str(i+1)+"行學號有誤\n")
                    if j==1:
                        Identity=data
                    if j==2:
                        result=0
                        for z in range (len(data)):
                            if  data[z]>= u'\u4e00' and data[z]<=u'\u9fff':
                                result=result
                            else:
                                result=result+1
                        if result>0:
                            Name="Error"
                            Errorcode=Errorcode+("第"+str(i+1)+"行姓名非中文\n")
                        else:
                            Name=data
                    if j==4:
                        if len(data)==5 or data=="unregistered":
                            CardID=data
                        else:
                            CardID="Error"
                            Errorcode=Errorcode+("第"+str(i+1)+"行卡號有誤\n")
                alldata[i]={
                         'SchoolID':SchoolID,
                         'CardID':CardID,
                         'Identity':Identity,
                         'Name':Name
                         }   
            if Errorcode!="":
                self.Error(Errorcode)
            else:
                check=self.Check("確認更新", "是否確認要上傳更新")
                if check==16384:
                    xqt.teadataupdate(DB, Collection, alldata)
                    self.ui.stuupedit.setText("")
                    self.teabutton()
    def teasetspan(self,data):
        self.ui.teaclearspan()
        additional_card=xqt.datadoublepath(data)
        if additional_card!=[]:
            if int(len(additional_card)/3)>1:
                for i in range(int(len(additional_card)/3)):
                    row=int(2*i+int(len(additional_card)/3)+1)
                    self.ui.teasetspan(int(additional_card[row]),0,int(additional_card[row-1]),1)
                    self.ui.teasetspan(int(additional_card[row]),1,int(additional_card[row-1]),1)
                    self.ui.teasetspan(int(additional_card[row]),2,int(additional_card[row-1]),1)
                    self.ui.teasetspan(int(additional_card[row]),3,int(additional_card[row-1]),1)
                    self.ui.teasetspan(int(additional_card[row]),5,int(additional_card[row-1]),1)
                    self.ui.teasetspan(int(additional_card[row]),6,int(additional_card[row-1]),1)
                    self.ui.teasetspan(int(additional_card[row]),7,int(additional_card[row-1]),1)
                    self.ui.teasetspan(int(additional_card[row]),8,int(additional_card[row-1]),1)
                    self.ui.teasetspan(int(additional_card[row]),9,int(additional_card[row-1]),1)
            else:
                self.ui.teasetspan(int(additional_card[2]),0,int(additional_card[1]),1)
                self.ui.teasetspan(int(additional_card[2]),1,int(additional_card[1]),1)
                self.ui.teasetspan(int(additional_card[2]),2,int(additional_card[1]),1)
                self.ui.teasetspan(int(additional_card[2]),3,int(additional_card[1]),1)
                self.ui.teasetspan(int(additional_card[2]),5,int(additional_card[1]),1)
                self.ui.teasetspan(int(additional_card[2]),6,int(additional_card[1]),1)
                self.ui.teasetspan(int(additional_card[2]),7,int(additional_card[1]),1)
                self.ui.teasetspan(int(additional_card[2]),8,int(additional_card[1]),1)
                self.ui.teasetspan(int(additional_card[2]),9,int(additional_card[1]),1)
    def teasubfunction(self):
        tabledata=self.ui.teatab.currentIndex().row()
        model = self.ui.teatab.model()
        if tabledata!=-1:
            CardID=str(model.data(model.index(tabledata,0)))
            sub.subfunction(DB, Collection, CollectionMAC, CardID)  
            self.teabutton()
        else:
            self.Error("請先選擇要編輯的資料")
#########學生資料############
    def stuupdata_grade(self):
        if self.ui.stuedit_2.text()!="":
            Errorcode=""
            row = self.ui.stutab.rowCount()
            alldata=[i for i in range(row)]
            for i in range(int(row)):
                for j in range(10):
                    item = self.ui.stutab.item(i, j)
                    data = str(item.text())
                    if j==0:
                        if len(data)==8 or data=="unregistered":
                            SchoolID=data
                        else:
                            SchoolID="Error"
                            Errorcode=Errorcode+("第"+str(i+1)+"行學號有誤\n")
                    if j==1:
                        Identity=data
                    if j==2:
                        result=0
                        for z in range (len(data)):
                            if  data[z]>= u'\u4e00' and data[z]<=u'\u9fff':
                                result=result
                            else:
                                result=result+1
                        if result>0:
                            Name="Error"
                            Errorcode=Errorcode+("第"+str(i+1)+"行姓名非中文\n")
                        else:
                            Name=data

                    if j==3:
                        Gender=data
                    if j==4:
                        if len(data)==10 or data=="unregistered":
                            CardID=data
                        else:
                            CardID="Error"
                            Errorcode=Errorcode+("第"+str(i+1)+"行卡號有誤\n")
                    if j==5:
                        Class=data
                    if j==6:
                        Weight=float(data)
                    if j==7:
                        DailyShouldDrink=float(data)
                    if j==8:
                        Height=float(data)
                    if j==9:
                        Number=data
                alldata[i]={
                         'SchoolID':SchoolID,
                         'CardID':CardID,
                         'Class':Class,
                         'DailyShouldDrink':DailyShouldDrink,
                         'Identity':Identity,
                         'Name':Name,
                         'Gender':Gender,
                         'Height':Height,
                         'Weight':Weight,
                         "Number":Number}   
            if Errorcode!="":
                self.Error(Errorcode)
            else:
                check=self.Check("確認更新", "是否確認要上傳更新")
                if check==16384:
                    xqt.dataupdata_grade(DB, Collection, alldata)
                    xqt.dataupdate(DB,Collection,alldata)
                    self.stucom()
                    self.ui.stuedit_2.setText("")                   
    def stuupdata(self):
        if self.ui.stuupedit.text()!="":
            Errorcode=""
            row = self.ui.stutab.rowCount()
            alldata=[i for i in range(row)]
            for i in range(int(row)):
                for j in range(10):
                    item = self.ui.stutab.item(i, j)
                    data = str(item.text())
                    if j==0:
                        if len(data)==8 or data=="unregistered":
                            SchoolID=data
                        else:
                            SchoolID="Error"
                            Errorcode=Errorcode+("第"+str(i+1)+"行學號有誤\n")
                    if j==1:
                        Identity=data
                    if j==2:
                        result=0
                        for z in range (len(data)):
                            if  data[z]>= u'\u4e00' and data[z]<=u'\u9fff':
                                result=result
                            else:
                                result=result+1
                        if result>0:
                            Name="Error"
                            Errorcode=Errorcode+("第"+str(i+1)+"行姓名非中文\n")
                        else:
                            Name=data
                    if j==3:
                        Gender=data
                    if j==4:
                        if len(data)==10 or data=="unregistered" :
                            CardID=data
                        else:
                            CardID="Error"
                            Errorcode=Errorcode+("第"+str(i+1)+"行卡號有誤\n")
                    if j==5:
                        Class=data
                    if j==6:
                        Weight=float(data)
                    if j==7:
                        DailyShouldDrink=float(data)
                    if j==8:
                        Height=float(data)
                    if j==9:
                        Number=data
                alldata[i]={
                         'SchoolID':SchoolID,
                         'CardID':CardID,
                         'Class':Class,
                         'DailyShouldDrink':DailyShouldDrink,
                         'Identity':Identity,
                         'Name':Name,
                         'Gender':Gender,
                         'Height':Height,
                         'Weight':Weight,
                         "Number":Number}   
            
                
            if Errorcode!="":
                self.Error(Errorcode)
            else:
                check=self.Check("確認更新", "是否確認要上傳更新")
                if check==16384:
                    xqt.dataupdate(DB,Collection,alldata)
                    self.stucom()
                    self.ui.stuupedit.setText("")
    def stusearview(self):
        self.ui.stuupedit.setText("")
        self.ui.stuedit_2.setText("")
        filename=QFileDialog.getOpenFileNames(None,'Open file',"C:","xlsx (*.xlsx *.xls )")
        data=[i for i in range(len(filename[0]))]
        text=""
        for i in range(len(filename[0])):
            text=text+" "+filename[0][i]
        self.ui.stuupedit.setText(text)
        if  len(filename[0])==1:
            try :
                data[0]=pd.read_excel(str(filename[0][0]))
                df=pd.DataFrame(data[0])
                df_1=pd.DataFrame(data[0])
            except:
                self.Error("檔案格式有誤")
        elif len(filename[0])==2:
            for i in range(len(filename[0])):
                try :
                    data[i]=pd.read_excel(str(filename[0][i]))
                except:
                    i=len(filename[0])
                    self.Error("檔案格式有誤")
            df=pd.DataFrame(data[0])
            df_1=pd.DataFrame(data[1])
        elif len(filename[0])>2:
            self.Error("選取檔案不得超過兩個")
        try:
            alldata=xqt.studatamerge(df, df_1)
            error=1
        except:
            alldata="Error"
            error=0
        if alldata!="Error" and len(alldata)!=0 and error==1:
            self.ui.stuclearspan()
            self.ui.sturowreload(alldata)
        else:
            self.Error("檔案內容有誤")
            self.ui.stuupedit.setText("")
    def stusearview_2(self):
        self.ui.stuupedit.setText("")
        self.ui.stuedit_2.setText("")
        filename=QFileDialog.getOpenFileNames(None,'Open file',"C:","xlsx (*.xlsx *.xls )")
        data=[i for i in range(len(filename[0]))]
        text=""
        for i in range(len(filename[0])):
            text=text+" "+filename[0][i]
        self.ui.stuedit_2.setText(text)
        if  len(filename[0])==1:
            try :
                data[0]=pd.read_excel(str(filename[0][0]))
                df=pd.DataFrame(data[0])
                df_1=pd.DataFrame(data[0])
            except:
                self.Error("檔案格式有誤")
        elif len(filename[0])==2:
            for i in range(len(filename[0])):
                try :
                    data[i]=pd.read_excel(str(filename[0][i]))
                except:
                    i=len(filename[0])
                    self.Error("檔案格式有誤")
            df=pd.DataFrame(data[0])
            df_1=pd.DataFrame(data[1])
        elif len(filename[0])>2:
            self.Error("選取檔案不得超過兩個")
        
        try:
            alldata=xqt.studatamerge(df, df_1)
            error=1
        except:
            alldata="Error"
            error=0
        if alldata!="Error" and len(alldata)!=0 and error==1:
            self.ui.stuclearspan()
            self.ui.sturowreload(alldata)
        else:
            self.Error("檔案內容有誤")
            self.ui.stuedit_2.setText("")
    def studelet(self):
        tabledata=self.ui.stutab.currentIndex().row()
        model = self.ui.stutab.model()
        if tabledata!=-1:
            message="是否要刪除\n"+"學號:"+model.data(model.index(tabledata,0))+"\n班級:"+model.data(model.index(tabledata,5))+"\n座號:"+model.data(model.index(tabledata,9))+"\n姓名:"+model.data(model.index(tabledata,2))
            check=self.Check("確認刪除",message)   
            if check==16384:
                SchoolID=model.data(model.index(tabledata,0))
                CardID=model.data(model.index(tabledata,4))
                data_2=xqt.Membersdata(DB, Collection,Search={'SchoolID':SchoolID,'Identity':'Student','CardID':CardID})
                if data_2!=False:
                    selector = { "SchoolID": str(model.data(model.index(tabledata,0))),"Identity":"Student" }
                    data2 = { "$set": { "Identity":"Alumni" } }
                    xqt.updateInDB(DB,Collection,selector, data2)
                try:
                    data=xqt.Membersdata(DB, CollectionMAC,Search={'SchoolID':SchoolID,'additionalCard':CardID})
                    error=0
                except:
                    error=1
                if error==0 and data!=False:
                    num=int(len(xqt.Membersdata(DB, CollectionMAC,Search={'SchoolID':SchoolID})))
                    if num==2:
                        print(model.data(model.index(tabledata,0)))
                        Condition={"SchoolID":model.data(model.index(tabledata,0))}
                        RJ.SC_DeleteDataFromDB(DB,CollectionMAC,Condition,"Many")
                    elif num!=2 and data_2!=False:
                        Condition={"SchoolID":model.data(model.index(tabledata,0))}
                        RJ.SC_DeleteDataFromDB(DB,CollectionMAC,Condition,"Many")
                    else:
                        Condition={"SchoolID":model.data(model.index(tabledata,0)),"additionalCard":model.data(model.index(tabledata,4))}
                        RJ.SC_DeleteDataFromDB(DB,CollectionMAC,Condition,"One")
                self.stucom()
        else:
            self.Error("請先選擇要刪除的資料")
    def stusearch(self):
        if self.ui.stusearedit.text()!="":
            SchoolID=self.ui.stusearedit.text()
            data=xqt.Membersdata(DB, Collection,Search={'SchoolID':SchoolID,'Identity':'Student'})
            if data==False:
                self.ui.stusearedit.setText("")
                self.Error("查無此學號")
            else:
                additional_card=xqt.Membersdata(DB, CollectionMAC,Search={'SchoolID':SchoolID})
                if additional_card==False:
                    self.ui.stusearedit.setText("")
                    self.ui.sturowreload(data)
                else:
                    for i in range(1,len(additional_card)):
                        insertdata={
                            'SchoolID':data[0]['SchoolID'],
                            'Identity':data[0]['Identity'],
                            'Name':data[0]['Name'],
                            'Gender':data[0]['Gender'],
                            'CardID':additional_card[i]['additionalCard'],
                            'Class':data[0]['Class'],
                            'Weight':data[0]['Weight'],
                            'DailyShouldDrink':data[0]['DailyShouldDrink'],
                            'Height':data[0]['Height'],
                            'Number':data[0]['Number']
                        }
                        data.insert( i, insertdata)
                    self.ui.stusearedit.setText("")
                    self.ui.sturowreload(data)
                    self.stusetspan(data)
        else:
            self.stucom()
        self.ui.stuupedit.setText("")
        self.ui.stuedit_2.setText("")
    
    def stucom(self):
        self.ui.stuedit_2.setText("")
        self.ui.stuupedit.setText("")
        text=str(self.ui.stucom.currentText())
        text_2=str(self.ui.stucom_2.currentText())
        if text=="All":
            self.ui.stucom_2.setEnabled(False)
            data=xqt.Membersdata_sequence(DB, Collection, "701")
            data_2=xqt.Membersdata_sequence(DB, Collection, "702")
            data_3=xqt.Membersdata_sequence(DB, Collection, "703")
            data_4=xqt.Membersdata_sequence(DB, Collection, "704")
            data_5=xqt.Membersdata_sequence(DB, Collection, "801")
            data_6=xqt.Membersdata_sequence(DB, Collection, "802")
            data_7=xqt.Membersdata_sequence(DB, Collection, "803")
            data_8=xqt.Membersdata_sequence(DB, Collection, "804")
            data_9=xqt.Membersdata_sequence(DB, Collection, "901")
            data_10=xqt.Membersdata_sequence(DB, Collection, "902")
            data_11=xqt.Membersdata_sequence(DB, Collection, "903")
            data_12=xqt.Membersdata_sequence(DB, Collection, "904")
            length=len(data)+len(data_2)+len(data_3)+len(data_4)+len(data_5)+len(data_6)+len(data_7)+len(data_8)+len(data_9)+len(data_10)+len(data_11)+len(data_12)
            alldata=[i for i in range(length)]
            length=0
            length_1=len(data)
            length_2=length_1+len(data_2)
            length_3=length_2+len(data_3)
            length_4=length_3+len(data_4)
            length_5=length_4+len(data_5)
            length_6=length_5+len(data_6)
            length_7=length_6+len(data_7)
            length_8=length_7+len(data_8)
            length_9=length_8+len(data_9)
            length_10=length_9+len(data_10)
            length_11=length_10+len(data_11)
            def one(alldata,data,length):
                for i in range(len(data)):
                    alldata[i+length]=data[i]
            def two(alldata,data,length):
                for i in range(len(data)):
                    alldata[i+length]=data[i]
            def three(alldata,data,length):
                for i in range(len(data)):
                    alldata[i+length]=data[i]
            def four(alldata,data,length):
                for i in range(len(data)):
                    alldata[i+length]=data[i]
            def five(alldata,data,length):
                for i in range(len(data)):
                    alldata[i+length]=data[i]
            def six(alldata,data,length):
                for i in range(len(data)):
                    alldata[i+length]=data[i]
            def seven(alldata,data,length):
                for i in range(len(data)):
                    alldata[i+length]=data[i]
            def eight(alldata,data,length):
                for i in range(len(data)):
                    alldata[i+length]=data[i]
            def night(alldata,data,length):
                for i in range(len(data)):
                    alldata[i+length]=data[i]
            def ten(alldata,data,length):
                for i in range(len(data)):
                    alldata[i+length]=data[i]
            def eleven(alldata,data,length):
                for i in range(len(data)):
                    alldata[i+length]=data[i]
            def tweleve(alldata,data,length):
                for i in range(len(data)):
                    alldata[i+length]=data[i]
            threads = []
            threads.append(threading.Thread(target = one(alldata,data,length), args = ()))        
            threads.append(threading.Thread(target = two(alldata,data_2,length_1), args = ()))
            threads.append(threading.Thread(target = three(alldata,data_3,length_2), args = ()))
            threads.append(threading.Thread(target = four(alldata,data_4,length_3), args = ()))
            threads.append(threading.Thread(target = five(alldata,data_5,length_4), args = ()))
            threads.append(threading.Thread(target = six(alldata,data_6,length_5), args = ()))
            threads.append(threading.Thread(target = seven(alldata,data_7,length_6), args = ()))        
            threads.append(threading.Thread(target = eight(alldata,data_8,length_7), args = ()))
            threads.append(threading.Thread(target = night(alldata,data_9,length_8), args = ()))
            threads.append(threading.Thread(target = ten(alldata,data_10,length_9), args = ()))
            threads.append(threading.Thread(target = eleven(alldata,data_11,length_10), args = ()))
            threads.append(threading.Thread(target = tweleve(alldata,data_12,length_11), args = ()))
            for i in range(len(threads)):
                threads[i].start()

            self.ui.sturowreload(alldata)
            self.stusetspan(alldata)
        else :
            self.ui.stucom_2.setEnabled(True)
            if text=="七年級":
                if text_2=="All":
                    data=xqt.Membersdata_sequence(DB, Collection, "701")
                    data_2=xqt.Membersdata_sequence(DB, Collection, "702")
                    data_3=xqt.Membersdata_sequence(DB, Collection, "703")
                    data_4=xqt.Membersdata_sequence(DB, Collection, "704")
                    length=len(data)+len(data_2)+len(data_3)+len(data_4)
                    alldata=[i for i in range(length)]
                    for i in range(len(data)):
                        alldata[i]=data[i]
                
                    for i in range(len(data_2)):
                        alldata[i+len(data)]=data_2[i]
                    for i in range(len(data_3)):
                        alldata[i+len(data)+len(data_2)]=data_3[i]
                    for i in range(len(data_4)):
                        alldata[i+len(data)+len(data_2)+len(data_3)]=data_4[i]
                    self.ui.sturowreload(alldata)
                    self.stusetspan(alldata)
                if text_2=="一班":
                    data=xqt.Membersdata_sequence(DB, Collection, "701")
                    self.ui.sturowreload(data)
                    self.stusetspan(data)
                if text_2=="二班":
                    data=xqt.Membersdata_sequence(DB, Collection, "702")
                    self.ui.sturowreload(data)
                    self.stusetspan(data)
                if text_2=="三班":
                    data=xqt.Membersdata_sequence(DB, Collection, "703")
                    self.ui.sturowreload(data)
                    self.stusetspan(data)
                if text_2=="四班":
                    data=xqt.Membersdata_sequence(DB, Collection, "704")
                    self.ui.sturowreload(data)
                    self.stusetspan(data)
            if text=="八年級":
                if text_2=="All":
                    data=xqt.Membersdata_sequence(DB, Collection, "801")
                    data_2=xqt.Membersdata_sequence(DB, Collection, "802")
                    data_3=xqt.Membersdata_sequence(DB, Collection, "803")
                    data_4=xqt.Membersdata_sequence(DB, Collection, "804")
                    length=len(data)+len(data_2)+len(data_3)+len(data_4)
                    alldata=[i for i in range(length)]
                    for i in range(len(data)):
                        alldata[i]=data[i]
                    length=len(data)
                    for i in range(len(data_2)):
                        alldata[i+length]=data_2[i]
                    length=length+len(data_2)
                    for i in range(len(data_3)):
                        alldata[i+length]=data_3[i]
                    length=length+len(data_3)
                    for i in range(len(data_4)):
                        alldata[i+length]=data_4[i]
                    self.ui.sturowreload(alldata)
                    self.stusetspan(alldata)
                if text_2=="一班":
                    data=xqt.Membersdata_sequence(DB, Collection, "801")
                    self.ui.sturowreload(data)
                    self.stusetspan(data)
                if text_2=="二班":
                    data=xqt.Membersdata_sequence(DB, Collection, "802")
                    self.ui.sturowreload(data)
                    self.stusetspan(data)
                if text_2=="三班":
                    data=xqt.Membersdata_sequence(DB, Collection, "803")
                    self.ui.sturowreload(data)
                    self.stusetspan(data)
                if text_2=="四班":
                    data=xqt.Membersdata_sequence(DB, Collection, "804")
                    self.ui.sturowreload(data)
                    self.stusetspan(data)
            if text=="九年級":
                if text_2=="All":
                    data=xqt.Membersdata_sequence(DB, Collection, "901")
                    data_2=xqt.Membersdata_sequence(DB, Collection, "902")
                    data_3=xqt.Membersdata_sequence(DB, Collection, "903")
                    data_4=xqt.Membersdata_sequence(DB, Collection, "904")
                    length=len(data)+len(data_2)+len(data_3)+len(data_4)
                    alldata=[i for i in range(length)]
                    for i in range(len(data)):
                        alldata[i]=data[i]
                
                    for i in range(len(data_2)):
                        alldata[i+len(data)]=data_2[i]
                    for i in range(len(data_3)):
                        alldata[i+len(data)+len(data_2)]=data_3[i]
                    for i in range(len(data_4)):
                        alldata[i+len(data)+len(data_2)+len(data_3)]=data_4[i]
                    self.ui.sturowreload(alldata)
                    self.stusetspan(alldata)
                if text_2=="一班":
                    data=xqt.Membersdata_sequence(DB, Collection, "901")
                    self.ui.sturowreload(data)
                    self.stusetspan(data)
                if text_2=="二班":
                    data=xqt.Membersdata_sequence(DB, Collection, "902")
                    self.ui.sturowreload(data)
                    self.stusetspan(data)
                if text_2=="三班":
                    data=xqt.Membersdata_sequence(DB, Collection, "903")
                    self.ui.sturowreload(data)
                    self.stusetspan(data)
                if text_2=="四班":
                    data=xqt.Membersdata_sequence(DB, Collection, "904")
                    self.ui.sturowreload(data)
                    self.stusetspan(data)
    def opensecond(self,data):
        self.dialog=SecondWindow()
        self.dialog.show()
        if data=="":
            self.dialog.setup_add()
        else:
            self.dialog.setup_edit(data)
    def stuedit(self):
        tabledata=self.ui.stutab.currentIndex().row()
        model = self.ui.stutab.model()
        if tabledata!=-1:
            SchoolID=model.data(model.index(tabledata,0))
            Name=model.data(model.index(tabledata,2))
            Gender=model.data(model.index(tabledata,3))
            CardID=model.data(model.index(tabledata,4))
            Class=model.data(model.index(tabledata,5))
            Weight=model.data(model.index(tabledata,6))
            Height=model.data(model.index(tabledata,8))
            Number=model.data(model.index(tabledata,9))
            data={
                'lineEdit':SchoolID,
                'lineEdit_2':Name,
                'lineEdit_3':CardID,
                'lineEdit_4':Class,
                'lineEdit_5':Weight,
                'lineEdit_6':Height,
                'lineEdit_7':Number,
                'lineEdit_8':Gender,
                'Check':'edit'
                }
            self.opensecond(data)
        else:
            self.Error("請先選擇要編輯的資料")
    def stuadd(self):
        self.opensecond("")
    def stusetspan(self,data):
        self.ui.stuclearspan()
        additional_card=xqt.datadoublepath(data)
        if additional_card!=[]:
            if int(len(additional_card)/3)>1:
                for i in range(int(len(additional_card)/3)):
                    row=int(2*i+int(len(additional_card)/3)+1)
                    self.ui.stusetspan(int(additional_card[row]),0,int(additional_card[row-1]),1)
                    self.ui.stusetspan(int(additional_card[row]),1,int(additional_card[row-1]),1)
                    self.ui.stusetspan(int(additional_card[row]),2,int(additional_card[row-1]),1)
                    self.ui.stusetspan(int(additional_card[row]),3,int(additional_card[row-1]),1)
                    self.ui.stusetspan(int(additional_card[row]),5,int(additional_card[row-1]),1)
                    self.ui.stusetspan(int(additional_card[row]),6,int(additional_card[row-1]),1)
                    self.ui.stusetspan(int(additional_card[row]),7,int(additional_card[row-1]),1)
                    self.ui.stusetspan(int(additional_card[row]),8,int(additional_card[row-1]),1)
                    self.ui.stusetspan(int(additional_card[row]),9,int(additional_card[row-1]),1)
            else:
                self.ui.stusetspan(int(additional_card[2]),0,int(additional_card[1]),1)
                self.ui.stusetspan(int(additional_card[2]),1,int(additional_card[1]),1)
                self.ui.stusetspan(int(additional_card[2]),2,int(additional_card[1]),1)
                self.ui.stusetspan(int(additional_card[2]),3,int(additional_card[1]),1)
                self.ui.stusetspan(int(additional_card[2]),5,int(additional_card[1]),1)
                self.ui.stusetspan(int(additional_card[2]),6,int(additional_card[1]),1)
                self.ui.stusetspan(int(additional_card[2]),7,int(additional_card[1]),1)
                self.ui.stusetspan(int(additional_card[2]),8,int(additional_card[1]),1)
                self.ui.stusetspan(int(additional_card[2]),9,int(additional_card[1]),1)
    def stusubfuction(self):
        tabledata=self.ui.stutab.currentIndex().row()
        model = self.ui.stutab.model()
        if tabledata!=-1:
            CardID=str(model.data(model.index(tabledata,0)))
            sub.subfunction(DB, Collection, CollectionMAC, CardID)  
            self.stucom()
        else:
            self.Error("請先選擇要編輯的資料")
#########水量報表##########
    def starbutton(self):
        def settext():
            self.ui.startdate.setText(self.ui.startdatewidge.selectedDate().toString("yyyy-MM-dd"))
            self.ui.startdatewidge.setVisible(False)
        
        if self.ui.startdatewidge.isVisible() == False:
            if self.ui.enddatewidge.isVisible() ==True:
                self.ui.enddatewidge.setVisible(False)
            self.ui.startdatewidge.setVisible(True)
            self.ui.startdatewidge.clicked.connect(settext)
        else:
            self.ui.startdatewidge.setVisible(False)
    def endbutton(self):
        def settext():
            self.ui.enddate.setText(self.ui.enddatewidge.selectedDate().toString("yyyy-MM-dd"))
            self.ui.enddatewidge.setVisible(False)
        
        if self.ui.enddatewidge.isVisible() == False:
            if self.ui.startdatewidge.isVisible() ==True:
                self.ui.startdatewidge.setVisible(False)
            self.ui.enddatewidge.setVisible(True)
            self.ui.enddatewidge.clicked.connect(settext)
        else:
            self.ui.enddatewidge.setVisible(False)
    def searchexcelpath(self):
        filename,_ = QFileDialog.getSaveFileName(None,  "文件保存","C:","xlsx (*.xlsx *.xls )")
        if not filename:
            self.ui.excelpath.setText("F:/")
            
        else :
            self.ui.excelpath.setText(str(filename))   
    def Calculate0(self):
        
        errorCode0 = ""
        error_cnt0 = 0
        
        StartDate0 = self.ui.startdate.text()
        EndDate0 = self.ui.enddate.text()
        SavePath0 = self.ui.excelpath.text()
        try:
            Threshold0 = int(self.ui.dataedit.text())
        except:
            Threshold0=-1
        print("=================")
        print(StartDate0)
        print(EndDate0)
        print(SavePath0)
        print(Threshold0)
        
        OldestTime0 = time.strptime("2021-10-25", "%Y-%m-%d")
        OldestTime0 = int(time.mktime(OldestTime0)) # 轉換為時間戳
        D_StartDate0 = time.strptime(StartDate0, "%Y-%m-%d")
        D_StartDate0 = int(time.mktime(D_StartDate0)) # 轉換為時間戳
        
        Str_Yesterday0 = RJ.Get_date_of_yesterday(display="%Y-%m-%d")
        Yesterday0 = time.strptime(Str_Yesterday0, "%Y-%m-%d")
        Yesterday0 = int(time.mktime(Yesterday0)) # 轉換為時間戳
        D_EndDate0 = time.strptime(EndDate0, "%Y-%m-%d")
        D_EndDate0 = int(time.mktime(D_EndDate0)) # 轉換為時間戳
        
        if OldestTime0-D_StartDate0>0 :
            errorCode0 = errorCode0 + "A" 
            error_cnt0 = error_cnt0 + 1
        
        if D_EndDate0-Yesterday0 > 0 :
            errorCode0 = errorCode0 + "B" 
            error_cnt0 = error_cnt0 + 1
        if SavePath0 == "F:/":
            errorCode0 = errorCode0 + "C" 
            error_cnt0 = error_cnt0 + 1
        if Threshold0 <=0:
            errorCode0 = errorCode0 + "D" 
            error_cnt0 = error_cnt0 + 1
        if D_StartDate0-D_EndDate0 >0:
            errorCode0 = errorCode0 + "E" 
            error_cnt0 = error_cnt0 + 1
        if error_cnt0>0:
            errmsg0 = "發生以下錯誤："
            errdisplay_cnt0 = 0
            if "A" in errorCode0:
                errdisplay_cnt0 = errdisplay_cnt0 + 1
                errmsg0 = errmsg0 + "\n" + str(errdisplay_cnt0) + ". " + "起始日期不得早於2021-10-25"
            if "B" in errorCode0:
                errdisplay_cnt0 = errdisplay_cnt0 + 1
                errmsg0 = errmsg0 + "\n" + str(errdisplay_cnt0) + ". " + "結束日期不得晚於" + Str_Yesterday0 + "(即昨天)"
            if "C" in errorCode0:
                errdisplay_cnt0 = errdisplay_cnt0 + 1
                errmsg0 = errmsg0 + "\n" + str(errdisplay_cnt0) + ". " + "儲存路徑不得為空"
            if "D" in errorCode0:
                errdisplay_cnt0 = errdisplay_cnt0 + 1
                errmsg0 = errmsg0 + "\n" + str(errdisplay_cnt0) + ". " + "採閥值須為正整數"
            if "E" in errorCode0:
                errdisplay_cnt0 = errdisplay_cnt0 + 1
                errmsg0 = errmsg0 + "\n" + str(errdisplay_cnt0) + ". " + "結束日期不得早於起始日期"
            # 迅速建立一個錯誤通知視窗
            self.Error(errmsg0)
        
        else:
          #  tkinter.Label(Tab0, text = "\n", fg="red").grid(column = 0, row = 5, sticky = 'EWNS') # 裝飾用，因為還沒開始拉Module進來用
            # 在這邊使用從其他自己寫的Module的程式，用來來產生圖表
            XinxingPG.excelFileAllData(From=StartDate0, To=EndDate0, SavePath=SavePath0, Threshold=Threshold0)
        
        print("=================")

   