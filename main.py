from tkinter import *
import smtplib
import speech_recognition as sr
from gtts import gTTS
import os
import pandas as pd
import xlsxwriter
import unittest

r=sr.Recognizer()
gpame=""
gage=""

with sr.Microphone() as source:
    speech = gTTS("Give Patient name")
    speech.save("patient.mp3")
    os.system("start patient.mp3")
    print("Give Patient name")
    audio=r.listen(source)

    try:
        pname=r.recognize_google(audio)
        print("Patient name is",pname)
        gpame=pname
    except:
        print("Sorry Plz say something not detectable...........")


    iage="Hii "+pname+".Please tell me your age and gender"
    speech = gTTS(iage)
    speech.save("age.mp3")
    os.system("start age.mp3")
    print("Hii" + pname + ".Please tell me your age and gender")
    ageaudio = r.listen(source)

    try:
        page = r.recognize_google(ageaudio)
        age=pname+"Your age is  "+page+"years .Thanks for sharing plz wait......."
        print("Your age is", page)
        gage=page
    except:
        print("Sorry Plz say something not detectable...........")

    doc="hii Doctor.Suggest the medications"
    speech = gTTS(doc)
    speech.save("document.mp3")
    os.system("start document.mp3")
    print("say yes for proceeding..")
    doc_ans = r.listen(source)

    try:
        ans = r.recognize_google(doc_ans)
        print(ans)
    except:
        print("Sorry Plz say something not detectable...........")
    if(ans=="yes" or ans=="Yes" or ans==None):
        doc = "tell me names medicines"
        speech = gTTS(doc)
        speech.save("doc.mp3")
        os.system("start doc.mp3")
        print("tell me names medicines")
        med = r.listen(source)

        try:
            mname = r.recognize_google(med)
            print(mname)
        except:
            print("Sorry cann't recognize your voice")
        doc2 = "Doctor please tell me time for consumption of medicines"
        speech = gTTS(doc2)
        speech.save("doc2.mp3")
        os.system("start doc2.mp3")
        print("Doctor please tell me time for consumption of medicines")
        medtime = r.listen(source)

        try:
            mtime = r.recognize_google(medtime)
            print(mtime)
        except:
            print("Sorry cann't recognize your voice")






root = Tk()
heading=Label(root,text="Medical Receipt",padx=300,pady=50,width=30,font=(None,30)).grid(row=1,column=2)
name=Label(root,text="Name of patient",font=(None,15),padx=20).grid(row=3,column=1)
iname=Entry( root ,width=50,bg="#f7f7f5",borderwidth=2)
iname.insert(0,pname)
iname.grid(row=3,column=2)


age=Label(root,text="Age and Gender of patient",font=(None,15),padx=20).grid(row=4,column=1)
iage=Entry( root ,width=50,bg="#f7f7f5",borderwidth=2)
iage.insert(0,page)
iage.grid(row=4,column=2)


medi=Label(root,text="1.Medicine name",font=(None,15),padx=20).grid(row=5,column=1)
imed=Entry( root ,width=50,bg="#f7f7f5",borderwidth=2)
imed.insert(0,mname)
imed.grid(row=5,column=2)
itime=Entry( root ,width=50,bg="#f7f7f5",borderwidth=2)
itime.insert(0,mtime)
itime.grid(row=6,column=2)


def printSomething():
    df = pd.DataFrame({
        'Name of patient': [pname],

        'Age and Gender': [page],
        'Medicine name': [mname],
        'Medicine dose':[mtime]})


    writer = pd.ExcelWriter('hospital.xlsx', engine='xlsxwriter')


    df.to_excel(writer, sheet_name='Sheet1')


    writer.save()
    print("Don't forgot to fill email address:")
    sender_email = "Please enter your name here"
    rec_email = email.get()
    password = "Please enter your password of email"
    message = "Patient Name ="+pname+"\n"+"Age and gender="+page+"\n"+"medicine="+mname+"\n"+"Medicine Time="+mtime
    server = smtplib.SMTP("smtp.gmail.com", 587)
    server.starttls()
    server.login(sender_email, password)
    print("login successfull")
    server.sendmail(sender_email, rec_email, message)
    print("email send successfully")



Patient_Email=Label(root,text="Email address of patient",font=(None,15),padx=20).grid(row=8,column=1)
email=Entry( root ,width=50,bg="#f7f7f5",borderwidth=2)
email.insert(0,"@gmail.com")
email.grid(row=8,column=2)

def test_name(pname):
        message="Patient name is not recognised"
        unittest.TestCase.assertIsNone(pname, message)

def test_gender(iage):
        
        message="Patient age is not recognised"
        unittest.TestCase.assertIsNone(iage, message)

test_name(gname);
test_age(gage);

mybutton=Button(root,text="Send email",padx=20,pady=7,command=printSomething,fg="black",bg="blue").grid(row=11,column=2)

root.mainloop()


