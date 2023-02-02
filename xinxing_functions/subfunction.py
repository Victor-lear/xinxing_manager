from XinxingQt import *

from tkinter import *
from tkinter import simpledialog
from tkinter import messagebox


def askCI():
    global a,Error,Error2
    root = Tk()
    root.withdraw()
    a = simpledialog.askstring(title='Card ID', prompt='Please Enter the  10 Digit Card ID of Your Additional Card:')
    if(a is not None):
        
        verify10(a)
        if Error2==0:
            Error=0
        else:
            Error=1
        return Error
    else:
        Error=0
        return Error



def verify10(data):
    global Error2
    try:
        int(data)
        Error2=1
    except:
        messagebox.showerror("showerror", 'Only numbers are allowed.')
        Error2=0
      #  popupRRA("How Would You Like to Proceed?")
    else:
        if(len(data)!=10 and len(data)!=5):
            messagebox.showerror("showerror", "Please enter 5 or 10 digits.")
            Error2=0
            #popupRRA("How Would You Like to Proceed?")

def register():
    pop.destroy()
    master = Tk()
    master.title('New User Registration')

    root  = Frame(master)
    root.pack(padx=10, pady=10, fill='x', expand=True)

    def getval():
        SchoolID = entry1.get()
        CardID = entry2.get()
        Class = entry3.get()
        DailyShouldDrink = entry4.get()
        Name = entry5.get()
        Gender = entry6.get()
        Height = entry7.get()
        Weight = entry8.get()
        Number = entry9.get()

        if(SchoolID != '' and CardID != '' and Class != '' and DailyShouldDrink != '' and Name != '' and Gender != '' and Height != '' and Weight != '' and Number != ''): 
            try:
                data = {'SchoolID': str(SchoolID),
                'CardID':str(CardID),
                'Class ': str(Class),
                'DailyShouldDrink': float(DailyShouldDrink),
                'Identity' : "Student" ,
                'Name ':str(Name),
                'Gender':str(Gender),
                'Height':float(Height),
                'Weight':float(Weight),
                "Number":str(Number)}
                if(ReadFromDB(db, collectionMD, SchoolID) == True):
                    messagebox.showerror('showerror', 'You are already registered.')
                else:
                    WriteInDB(db, collectionMD, data)
            except:
                messagebox.showerror('showerror', 'Please enter in all fields correctly.')
        else:
            messagebox.showerror('showerror', 'Please enter in all fields.')
        
        master.destroy()

    label1 = Label(root, text="School ID: ").grid(row=0)
    label2 = Label(root, text="Card ID: ").grid(row=1)
    label3 = Label(root, text="Class: ").grid(row=2)
    label4 = Label(root, text="DailyShouldDrink (mL): ").grid(row=3)
    label5 = Label(root, text="Name: ").grid(row=4)
    label6 = Label(root, text="Gender: ").grid(row=5)
    label7 = Label(root, text="Height: (cm)").grid(row=6)
    label8 = Label(root, text="Weight: (kg)").grid(row=7)
    label9 = Label(root, text="Number: ").grid(row=8)

    entry1 = Entry(root)
    entry2 = Entry(root)
    entry3 = Entry(root)
    entry4 = Entry(root)
    entry5 = Entry(root)
    entry6 = Entry(root)
    entry7 = Entry(root)
    entry8 = Entry(root)
    entry9 = Entry(root)

    entry1.grid(row=0, column=1)
    entry2.grid(row=1, column=1)
    entry3.grid(row=2, column=1)
    entry4.grid(row=3, column=1)
    entry5.grid(row=4, column=1)
    entry6.grid(row=5, column=1)
    entry7.grid(row=6, column=1)
    entry8.grid(row=7, column=1)
    entry9.grid(row=8, column=1)

    button = Button(root, text="Register", command=lambda: getval()).grid(row=9, column=1)
    master.mainloop()


def popupRRA(msg):
    global pop
    pop = Tk()
    pop.title(' ')
    pop.geometry('500x200')
    pop.configure(background='gray')
    
    label = Label(pop, text=msg, font=("Verdana", 24), background='gray', foreground='white')
    label.pack(side='top', fill='x', pady=50)

    REENTER = Button(pop, text='Re-enter card ID', font=("Verdana", 16), command=lambda: askCI())
    REENTER.pack(side='left', padx=30)

  #  REG = Button(pop, text='Register as a new user', font=("Verdana", 16), command=lambda: register())
   # REG.pack(side='right', padx=30)

    pop.mainloop()

def checkPrim(s1, a1):
    if(Membersdata(DB=db, Collection=collectionMAC, Search={'SchoolID': s1, 'primaryCard': a1}) == False):
        return False
    else:
        return True

def checkAdd(s1, a1):
    if(Membersdata(DB=db, Collection=collectionMAC, Search={'SchoolID': s1, 'additionalCard': a1}) == False):
        return False
    else:
        return True

def CountInDB(DB, Collection, Search={}):
    mongo_url_01 = "mongodb://admin:bmwee8097218@140.118.122.115:31383/"
    mongo_url_02 = "mongodb://admin:bmwee8097218@140.118.122.115:30415/"
    mongo_url_03 = "mongodb://admin:bmwee8097218@140.118.122.115:30708/"
    try:
        conn = MongoClient(mongo_url_01) 
        db = conn[DB]
        collection = db[Collection]
        return collection.find(Search).count()
    except:
        try:
            conn = MongoClient(mongo_url_02)
            db = conn[DB]
            collection = db[Collection]
            return collection.find(Search).count()
        except:
            try:
                conn = MongoClient(mongo_url_03)
                db = conn[DB]
                collection = db[Collection]
                return collection.find(Search).count()
            except:
                return -1

def DeleteInDB(DB,Collection,Search={}):
    mongo_url_01 = "mongodb://admin:bmwee8097218@140.118.122.115:31383/"
    mongo_url_02 = "mongodb://admin:bmwee8097218@140.118.122.115:30415/"
    mongo_url_03 = "mongodb://admin:bmwee8097218@140.118.122.115:30708/"
    try:
        conn = MongoClient(mongo_url_01) 
        db = conn[DB]
        collection = db[Collection]
        collection.delete_many(Search)
    except:
        try:
            conn = MongoClient(mongo_url_02)
            db = conn[DB]
            collection = db[Collection]
            collection.delete_many(Search)
        except:
            try:
                conn = MongoClient(mongo_url_03)
                db = conn[DB]
                collection = db[Collection]
                collection.delete_many(Search)
            except:
                print("no occurences found")
def subfunction(DB,Collection,CollectionMAC,CardID):
    global db, collectionMAC, collectionMD,s,a,pop,Error,Error2
    db = DB
    collectionMAC = CollectionMAC
    collectionMD = Collection
    s=CardID
    LARGE_FONT = ("Verdana", 24)
    NORM_FONT = ("Verdana", 16)
    askCI()
    if Error==1:
        if(checkPrim(s, a) ==  True):
            messagebox.showerror("showerror", "Sorry, this ID is already registered as your primary card.")
        elif(checkAdd(s, a) == True):
            messagebox.showerror("showerror", "Sorry, this ID is already registered as your additional card")
        else:
            root = Tk()
            
            root.withdraw()
            index = ""
            if(CountInDB(db, collectionMAC, {'SchoolID': s}) == 0):
                data = Membersdata(db, collectionMD, {'SchoolID': s})
                data = str(data)
                index = data.partition("'CardID': '")
                try:
                    INPUT = index[2][0:10]
                    int(INPUT)
                except:
                    INPUT = index[2][0:5]
                    int(INPUT)
                verify10(INPUT)
                WriteInDB(db, collectionMAC, {'SchoolID': s, 'primaryCard': INPUT, 'additionalCard': INPUT})
            else:
                data = Membersdata(db, collectionMAC, { "$expr": { "$and": [ { "$eq": [ "$SchoolID", s ]},
                             { "$eq": [ "$primaryCard", "$additionalCard" ]}]}})
                data = str(data)
                index = data.partition("'primaryCard': '")
                try:
                    INPUT = index[2][0:10]
                    int(INPUT)
                except:
                    INPUT = index[2][0:5]
                    int(INPUT)
            WriteInDB(db, collectionMAC, {'SchoolID': s, 'primaryCard': INPUT, 'additionalCard': a})
            DeleteInDB(db, collectionMD,{'SchoolID': 'unregistered', 'CardID': a})