from tkinter import*
from tkinter import messagebox,filedialog
from pygame import mixer # this will give usmusic
import speech_recognition
from email.message import EmailMessage
import os
import smtplib as s
import imghdr
import pandas

check=False
root=Tk()

def browse():
    global final_emails
    path=filedialog.askopenfilename(initialdir='c:/',title="Select Excel File")#this will give the whole path
    #now to get the imail adress from the excel we use panadas
    if path==None:#none menas null in python
        messagebox.showerror('Error','Please Select an Excel File')
    else:
        data=pandas.read_excel(path)#open py excel to read excel data
        if 'Email' in data.columns:
            emails=list(data['Email'])#converting to list
            print(emails)
            final_emails=[]
            for i in emails:
                if pandas.isnull(i)==False:
                    final_emails.append(i)
            if len(final_emails)==0:
                messagebox.showerror("Error",'File does not contain any email adress')   

            else:
                toEntryField.config(state=NORMAL)#beacuse hamne iski state disabled kr rkhi hai
                toEntryField.insert(0,os.path.basename(path))
                toEntryField.config(state="readonly")#beCAUE NAME NHI CHANGE KR SAKE VAPIS THATS HWY VAPIS READ ONLY KR DENGE
                totalLabel.config(text='Total-'+str(len(final_emails))) #to conver into the str beeacuse the concat krne ke liye string mandatory
                sentLabel.config(text='Sent:')
                leftLabel.config(text='Left:')
                failedLabel.config(text='Failed:')


def button_check():
    if choiceVariable.get()=='multiple':
        browseButton.config(state=NORMAL)#disbale se normal kr di
        toEntryField.config(state='readonly')
    else:
        if choiceVariable.get()=='single':
            browseButton.config(state=DISABLED)
            toEntryField.config(state=NORMAL)    


def attachment():
    global filename,filetype,file,check
    check=True
    file=filedialog.askopenfilename(initialdir='c:/',title='Select File')#this return the path
    filetype=file.split('.')
    filetype=filetype[1]
    filename=os.path.basename(file)  # this return the file name useing os
    # print(filename)
    textarea.insert(END,f'\n{filename}\n') 
 



def sendingEmail(toAdress,subject,body):
    f=open("credentials.txt",'r')
    for i in f:
        login_details=(i.split(','))
        # print(login_details[0])
        

    # ob=s.SMTP("smtp.gmail.com",587)#idar mode likhna hai agr  gmail hai to gmail,vrna reiffmail and port number
    # ob.starttls()
    # ob.login(login_details[0],login_details[1])
    # message=(f"{subject}\n\n{body}")
    # ob.sendmail(login_details[0],toAdress,message)
    # print("send succesfully")
    # first way with subject problem

    em=EmailMessage()
    em['subject']=subject
    em['to']=toAdress
    em['from']=login_details[0]
    em.set_content(body)
    if check:#menas if user doesnot select any attachment and this code not  exexute to check whter attachment is selected or not
        if filetype=='png' or filetype=='jpg' or filetype=='jpeg':
            f=open(file,'rb')
            file_data=f.read()
            subtype=imghdr.what(file)#this gives ht eextension of the file
            em.add_attachment(file_data,maintype='image',subtype=subtype,filename=filename)  

        else:
            f=open(file,'rb')
            file_data=f.read()
            em.add_attachment(file_data,maintype='application',subtype='oct-stream',filename=filename)
     

    ob=s.SMTP("smtp.gmail.com",587)#idar mode likhna hai agr  gmail hai to gmail,vrna reiffmail and port number
    ob.starttls()
    ob.login(login_details[0],login_details[1]) 
    ob.send_message(em)
    x=ob.ehlo()
    if x[0]==250:
        return 'sent'
    else:
        return 'failed'
    # messagebox.showinfo("Information","Email ")




def send_Email():
    if toEntryField.get()=='' or subjectEntryField.get()=='' or textarea.get(1.0,END)=='\n':
        messagebox.showerror("Error","All fields required")
        
    else:
        if choiceVariable.get()=='single':
            result=sendingEmail(toEntryField.get(),subjectEntryField.get(),textarea.get(1.0,END))
            if result=='sent':
                messagebox.showinfo("Infomation","Sent Successfully")
            if result=='failed':
                 messagebox.showerror("Email","Failed")    

    if choiceVariable.get()=='multiple':#finalemails ko globL BANYENGE BECAUSE ARGS PASS KRNA HAI
        sent=0
        failed=0
        for x in final_emails:
            result=sendingEmail(x,subjectEntryField.get(),textarea.get(1.0,END))
            if result=='sent':
                sent+=1
            if result=='failed':
                failed+=1

            totalLabel.config(text='')
            sentLabel.config(text='Sent:' + str(sent))
            leftLabel.config(text='Left:' + str(len(final_emails) - (sent + failed)))
            failedLabel.config(text='Failed:' + str(failed))

            totalLabel.update()
            sentLabel.update()
            leftLabel.update()
            failedLabel.update() 

        messagebox.showinfo("Email","Emails are sent succesfully")       

        



def speak():
    mixer.init()
    mixer.music.load('music1.mp3')
    mixer.music.play()
    sr=speech_recognition.Recognizer()#so we can acces the method of this classes
    with speech_recognition.Microphone() as m:
        try:
            sr.adjust_for_ambient_noise(m,duration=0.2)
            audio=sr.listen(m)
            text=sr.recognize_google(audio)
            textarea.insert(END,text+'.')
        except: 
            print("sorry") 


def iexit():
    #this retrun tru and false
    result=messagebox.askyesno("Notification","Do you want to exit?")
    if result:
        root.destroy()
    else:
        pass 

def clear():
    toEntryField.delete(0,END)
    subjectEntryField.delete(0,END)
    textarea.delete(1.0,END)#in this text area index starts with 1.01 


def settings():
    def clear1():
        fromEntryField.delete(0,END)
        passwordEntryField.delete(0,END)


    def save():
        if fromEntryField.get()=='' or passwordEntryField.get()=='':
            messagebox.showerror("Error","Please fill the details first",parent=root1)#by default konsi window khuli rhe hai to root1 khuli rhe 

        else:
            f=open('credentials.txt','w')
            f.write(fromEntryField.get()+','+passwordEntryField.get())
            f.close()
            messagebox.showinfo("Details","Details Saved",parent=root1)        

    root1=Toplevel()#to create another windoe we use toplevel
    root1.title("Setting")
    root1.geometry('500x340+350+90')#menas kitni x axis and yaxis se window open ho
    root1.config(bg='SlateBlue1')
    Label(root1,text="Login Details",image=logoImage,compound=LEFT,font=('goudy old style',40,'bold'),fg='white',bg='gray20').grid(row=0,column=0,padx=40,)

    fromLabelFrame=LabelFrame(root1,text='From (Email Address)',font=('times new roman',16,'bold'),bd=5,fg='white',bg='SlateBlue1')
    fromLabelFrame.grid(row=1,column=0,pady=20)

    fromEntryField=Entry(fromLabelFrame,font=('times new roman',18,'bold'),width=30)
    fromEntryField.grid(row=0,column=0)

    passwordLabelFrame=LabelFrame(root1,text='Password',font=('times new roman',16,'bold'),bd=5,fg="white",bg='SlateBlue1')
    passwordLabelFrame.grid(row=2,column=0,pady=20)

    passwordEntryField=Entry(passwordLabelFrame,font=('times new roman',18,'bold'),width=30,show="*")#for passowrd show only *
    # using show=*
    passwordEntryField.grid(row=0,column=0)

    Button(root1,text="SAVE",font=('times new roman',18,'bold'),cursor='hand2',bg='gold2',fg='black',command=save).place(x=110,y=280)
    Button(root1,text="CLEAR",font=('times new roman',18,'bold'),cursor='hand2',bg='gold2',fg='black',command=clear1).place(x=250,y=280)
    # f=open('credentials.txt','r')
    # for i in f:
    #     credentials=i.split(',')
    #     fromEntryField.insert(0,credentials[0])
    #     passwordEntryField.insert(0,credentials[1])
    # f.close() 




    root1.mainloop()    

root.title('Email Sender')
root.geometry('780x620+100+50')
#using this we cannot mincrease width and min
root.config(bg='SlateBlue1')
titleFrame=Frame(root,bg='white')
titleFrame.grid(row=0,column=0)
root.resizable(0,0)

logoImage=PhotoImage(file='email.png')

titleLabel=Label(titleFrame,text="   Email Sender",image=logoImage,compound=LEFT,font=('Goudy old Style',28,'bold'),bg="white",fg='SlateBlue1')
#if we want both imag and text the we have to use compund
titleLabel.grid(row=0,column=0)
settingImage=PhotoImage(file='setting.png')
# to remove baorder bd=0
# whenever i click on button it will give me grey color so use activebackgroun=white
Button(titleFrame,image=settingImage,bd=0,bg='white',cursor='hand2',command=settings).grid(row=0,column=1)

chooseframe=Frame(root,bg='SlateBlue1')
#after giving space in buttons white color of fram is shown thatswhy i give color to fram choosefram
chooseframe.grid(row=1,column=0,pady=10)
choiceVariable=StringVar()


singleRadioButton=Radiobutton(chooseframe,text='Single',font=('times new roman',25,'bold'),variable=choiceVariable,value='single',bg='SlateBlue1',activebackground='SlateBlue1',padx=20,command=button_check)#after cling the single  the coloe is whitw so thats why we sue activebg-blue

singleRadioButton.grid(row=0,column=0)

multipleRadioButton=Radiobutton(chooseframe,text='Multiple',font=('times new roman',25,'bold'),variable=choiceVariable,value='multiple',bg='SlateBlue1',activebackground='SlateBlue1',padx=20,command=button_check).grid(row=0,column=1)

choiceVariable.set('single')

tolabelFrame=LabelFrame(root,text='To (Email Adress)',font=('times new roman',16,'bold'),bd=5,fg='white',bg='SlateBlue1')
tolabelFrame.grid(row=2,column=0,padx=95)

toEntryField=Entry(tolabelFrame,font=('times new roman',18,'bold'),width=30)#for increase the size of font in entry  use
toEntryField.grid(row=0,column=0)#why row and col=0 becaus it is a first thing that i input in the tolabel fram

browseImage=PhotoImage(file='browse.png')

browseButton=Button(tolabelFrame,text="Browse",image=browseImage,compound=LEFT,font=("Arial",12,'bold'),bd=0,cursor='hand2',bg='SlateBlue1',activebackground='dodger blue2',state=DISABLED,command=browse)
browseButton.grid(row=0,column=1,padx=20)
#state disabled means disable rehega intially

subjectlabelFrame=LabelFrame(root,text='Subject',font=('times new roman',16,'bold'),bd=5,fg='white',bg='SlateBlue1')
subjectlabelFrame.grid(row=3,column=0)

subjectEntryField=Entry(subjectlabelFrame,font=('times new roman',18,'bold'),width=30)
subjectEntryField.grid(row=0,column=0)

emailLabelFrame=LabelFrame(root,text='Compose Email',font=('times new roman',16,'bold'),bd=5,fg='white',bg='SlateBlue1')
emailLabelFrame.grid(row=4,column=0,padx=20)

micImage=PhotoImage(file='mic.png')
attachImage=PhotoImage(file='attachments.png')
Button(emailLabelFrame,text="Speak",image=micImage,compound=LEFT,font=("Arial",12,'bold'),bd=0,cursor='hand2',bg='SlateBlue1',activebackground='SlateBlue1',command=speak).grid(row=0,column=0)
Button(emailLabelFrame,text="Attachment",image=attachImage,compound=LEFT,font=("Arial",12,'bold'),bd=0,cursor='hand2',bg='SlateBlue1',activebackground='SlateBlue1',command=attachment).grid(row=0,column=1)

textarea=Text(emailLabelFrame,font=("times new roman",14),height=9)#for fixing height we sue height and width
textarea.grid(row=1,column=0,columnspan=2)

sendImage=PhotoImage(file="send.png")
Button(root,image=sendImage,bd=0,bg='dodger blue2',cursor='hand2',activebackground='SlateBlue1',command=send_Email).place(x=490,y=545)

clearImage=PhotoImage(file="clear.png")
Button(root,image=clearImage,bd=0,bg='dodger blue2',cursor='hand2',activebackground='SlateBlue1',command=clear).place(x=590,y=550)

exitImage=PhotoImage(file="exit.png")
Button(root,image=exitImage,bd=0,bg='dodger blue2',cursor='hand2',activebackground='SlateBlue1',command=iexit).place(x=690,y=555)

totalLabel=Label(root,font=('times new roman',18,'bold'),bg='SlateBlue1',fg="black")
totalLabel.place(x=10,y=560)

sentLabel=Label(root,font=('times new roman',18,'bold'),bg='SlateBlue1',fg="black")
sentLabel.place(x=100,y=560)

leftLabel=Label(root,font=('times new roman',18,'bold'),bg='SlateBlue1',fg="black")
leftLabel.place(x=190,y=560)

failedLabel=Label(root,font=('times new roman',18,'bold'),bg='SlateBlue1',fg="black")
failedLabel.place(x=280,y=560)



root.mainloop()
