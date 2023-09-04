
from tkinter import *
from tkinter import messagebox,filedialog
from pygame import mixer
import speech_recognition
from email.message import EmailMessage
import smtplib
import os
import imghdr
import pandas
'''def sendingemail(toaddress,subject,body):
    f=open('credential.txt','r')
    tEntryField.get(),textArea.get(1.0,END))'''
check= False
def browse( ):
    global final_emails
    path=filedialog.askopenfilename(initialdir='c:/',title='Select Excel File')
    print(path)
    if path=='':
        messagebox.showerror('Error','Please select an Excel File')

    else:
        data = pandas.read_excel(path)
    print(data)
    # Specify the expected column name
    expected_column_name = 'Email'
    print(data.columns)
    expected_column_name = 'Email'

    # Strip whitespaces from column names
    data.columns = data.columns.str.strip()

    # Check if the expected column name is present in the stripped column names
    if expected_column_name in data.columns:
        emails = list(data[expected_column_name])
        print(emails)
        # Rest of the code remains the same
        final_emails = []
        for i in emails:
            if pandas.isnull(i) == False:
                final_emails.append(i)
        print(final_emails)
        print(len(final_emails))

        if len(final_emails) == 0:
            messagebox.showerror('Error', 'File does not contain any email addresses')
        else:
            toEntryField.config(state=NORMAL)
            toEntryField.insert(0, os.path.basename(path))
            toEntryField.config(state='readonly')
            totalLabel.config(text='Total:' + str(len(final_emails)))
            sentLabel.config(text='Sent:')
            leftLabel.config(text='Left:')
            failLabel.config(text='Failed:')
    else:
        # Handle the case when the column is not found
        messagebox.showerror('Error', f"File does not contain the column '{expected_column_name}'")

    # Check if the expected column name is present in the data columns










def button_check():
    if(choice.get()=='multiple'):
        browsebutton.config(state=NORMAL)
        toEntryField.config(state='readonly')
    if (choice.get() == 'single'):
        browsebutton.config(state=DISABLED)
        toEntryField.config(state=NORMAL)



def attachment():

   #filepath=filedialog.askopenfilename(initaildir='c:/',title='select file')

   global filename,filetype,filepath,check
   check = True
   filepath = filedialog.askopenfilename(initialdir='c:/', title='Select File')

   filetype=filepath.split('.')
   filetype=filetype[1]
   filename=os.path.basename(filepath)
   textArea.insert(END,f'\n{filename}\n')


def sendingemail(toaddress, subject, body):
    with open('credential.txt', 'r') as f:
        credentials = f.readline().strip().split(',')

    message = EmailMessage()
    message['subject'] = subject
    message['to'] = toaddress
    message['from'] = credentials[0]
    message.set_content(body)
    if check:
      if filetype=='png'or filetype=='jpg' or filetype=='jpeg':
        f=open(filepath,'rb')
        file_data=f.read()
        subtype=imghdr.what(filepath)
        message.add_attachment(file_data,maintype='Image',subtype=subtype,filename=filename)
      else:
        f = open(filepath, 'rb')
        file_data = f.read()
        message.add_attachment(file_data, maintype='application', subtype='octet-stream', filename=filename)
    else:
        pass
    s = smtplib.SMTP('smtp.gmail.com', 587)
    s.starttls()
    s.login(credentials[0], credentials[1])
    s.send_message(message)
    x=s.ehlo()
    if x[0]==250:
        return 'sent'
    else :
        return 'failed'
    s.quit()

    messagebox.showinfo('Information', 'Email sent successfully')


def sent_email():
    if toEntryField.get().strip() == '' or tosubjectEntryField.get().strip() == '' or textArea.get(1.0,
                                                                                                   END).strip() == '':
        messagebox.showerror('Error', 'All fields are required', parent=root)
    else:
        if choice.get() == 'single':
            result=sendingemail(toEntryField.get(), tosubjectEntryField.get(), textArea.get(1.0, END))
            if result=='sent':
                messagebox.showinfo('success','Email is sent successfuly')
            if result=='failed':
                messagebox.showerror('Error','Email is not sent')




        if choice.get() == 'multiple':
            sent=0
            failed=0
            for x in final_emails:
                result=sendingemail(x, tosubjectEntryField.get(), textArea.get(1.0, END))
                if result == 'sent':
                    messagebox.showinfo('success', 'Email is sent successfuly')
                    sent+=1
                if result == 'failed':
                    messagebox.showerror('Error', 'Email is not sent')
                    failed+=1
                totalLabel.config(text='' )
                sentLabel.config(text='Sent:'+str(sent))
                leftLabel.config(text='Left:'+str(len(final_emails)-(sent+failed)))
                failLabel.config(text='Failed:'+str(failed))
                totalLabel.update()
                sentLabel.update()
                leftLabel.update()
                failLabel.update()
            messagebox.showinfo('success','Emails are sent successfully')

def settings():
    def clear1():
        fromEntryField.delete(0,END)
        passwordEntryField.delete(0,END)

    def save():
       if passwordEntryField.get()==' 'and fromEntryField.get()==' ':
           messagebox.showerror('Error','All fields are required')
       else:
           f=open('credential.txt','w')
           f.write(fromEntryField.get()+','+passwordEntryField.get())
           f.close()
           messagebox.showinfo('Information','Credentials Saved Successfully',parent=root1)


    root1=Toplevel()
    root1.title('setting')
    root1.geometry('650x380+350+90')
    root1.configure(bg='light blue')
    Label(root1,text='Credentials Setting',image=logoImage,compound=LEFT,font=('goudy old style',40,'bold'),fg='white',bg='grey20').grid(padx=60)
    fromLabelframe = LabelFrame(root1, text='to (From Address)', font=('time new roman', 16, 'bold'), bd=5, fg='white',
                              bg='light blue')
    fromLabelframe.grid(row=1, column=0,pady=20)
    fromEntryField = Entry(fromLabelframe, font=('time new roman', 18, 'bold'), width=30)
    fromEntryField.grid(row=0, column=0, padx=5, pady=10)
    passwordLabelframe = LabelFrame(root1, text='Password', font=('time new roman', 16, 'bold'), bd=5, fg='white',
                                bg='light blue')
    passwordLabelframe.grid(row=2, column=0, pady=20)
    passwordEntryField = Entry(passwordLabelframe, font=('time new roman', 18, 'bold'), width=30,show='*')
    passwordEntryField.grid(row=0, column=0)
    Button(root1, text='SAVE',  font=('arial', 16, 'bold'),bg='gold2',fg='black',
           cursor='hand2', bd=0, activebackground='light blue',command=save).place(x=210,y=320)
    Button(root1, text='CLEAR', font=('arial', 16, 'bold'), bg='gold2', fg='black',
           cursor='hand2', bd=0, activebackground='light blue',command=clear1).place(x=340, y=320)
    f=open('credential.txt','r')
    for i in f:
        credentials=i.split(',')
    fromEntryField.insert(0,credentials[0])
    passwordEntryField.insert(0,credentials[1])


    root1.mainloop()
def iexit():
   result=messagebox.askyesno('notification','Do you want to exit?')
   if result:
       root.destroy()
   else:
       pass



def clear():
    toEntryField.delete(0,END)
    tosubjectEntryField.delete(0,END)
    textArea.delete(1.0,END)

def speak():
     mixer.init()
     mixer.music.load('music1.mp3')
     mixer.music.play()
     sr=speech_recognition.Recognizer()
     with speech_recognition.Microphone() as m:
         try:
             sr.adjust_for_ambient_noise(m,duration=0.2)
             audio=sr.listen(m)
             text=sr.recognize_google(audio)
             textArea.insert(END,text+'.')

         except:
             pass




root=Tk()
root.title('Mail Application')
root.geometry('780x620+100+50')
root.resizable(0,0)
root.configure(bg='light blue')
titleFrame=Frame(root,bg='white')
titleFrame.grid(row=0,column =0)
logoImage=PhotoImage(file='email2.png')
titleLabel=Label(titleFrame,text=' Email Sender',image=logoImage,compound=LEFT,font=('Goudy old style',28,'bold'),bg='white',fg='dodger blue2')

titleLabel.grid(row=0,column=0)
settingImage=PhotoImage(file='setting.png')
Button(titleFrame,image=settingImage,bd=0,bg='white',cursor='hand2',activebackground='white',command=settings).grid(row=0,column=1,padx=20)
chooseFrame=Frame(root,bg='light blue')
chooseFrame.grid(row=1,column=0,padx=20,pady=10)
choice=StringVar()


singleRadiobutton=Radiobutton(chooseFrame,text='Single',font=('time new roman',25,'bold'),variable=choice,value='single',bg='light blue',activebackground='light blue',command=button_check)
singleRadiobutton.grid(row=0,column=0)
multipleRadiobutton=Radiobutton(chooseFrame,text='Multiple',font=('time new roman',25,'bold'),variable=choice,value='multiple',bg='light blue',activebackground='light blue',command=button_check)
multipleRadiobutton.grid(row=0,column=1,padx=20)

choice.set('single')


toLabelframe=LabelFrame(root,text='to (Email Address)',font=('time new roman',16,'bold'),bd=5,fg='white',bg='light blue')
toLabelframe.grid(row=2,column=0,padx=100)
toEntryField=Entry(toLabelframe,font=('time new roman',18,'bold'),width=30)
toEntryField.grid(row=0,column=0,padx=5,pady=10)
browseImage=PhotoImage(file='browse.png')
browsebutton=Button(toLabelframe,text='browse',image=browseImage,compound=LEFT,font=('arial',12,'bold'),cursor='hand2',bd=0,activebackground='light blue',command=browse,state=DISABLED)
browsebutton.grid(row=0,column=1,padx=20)

tosubjectLabelframe=LabelFrame(root,text='Subject',font=('time new roman',16,'bold'),bd=5,fg='white',bg='light blue')
tosubjectLabelframe.grid(row=3,column=0,pady=10)
tosubjectEntryField=Entry(tosubjectLabelframe,font=('time new roman',18,'bold'),width=30)
tosubjectEntryField.grid(row=0,column=0)
toemailLabelframe=LabelFrame(root,text='Compose Email',font=('time new roman',16,'bold'),bd=5,fg='white',bg='light blue')
toemailLabelframe.grid(row=4,column=0)
micImage=PhotoImage(file='mic.png')
micbutton=Button(toemailLabelframe,text='Speak',image=micImage,compound=LEFT,font=('arial',12,'bold'),cursor='hand2',bd=0,activebackground='light blue',command=speak)
micbutton.grid(row=0,column=0,padx=20)
attachImage=PhotoImage(file='attachments.png')
attachbutton=Button(toemailLabelframe,text='Attachment',image=attachImage,compound=LEFT,font=('arial',12,'bold'),cursor='hand2',bd=0,activebackground='light blue',command=attachment)
attachbutton.grid(row=0,column=1,padx=20)
textArea=Text(toemailLabelframe,font=('new roman times',14,'bold'),height=8)

textArea.grid(row=1,column=0,columnspan=2)
sendImage=PhotoImage(file='send.png')
sentbutton=Button(root,image=sendImage,bd=0,bg='light blue',cursor='hand2',activebackground='light blue',command=sent_email)
sentbutton.place(x=490,y=550)


clearImage=PhotoImage(file='clear.png')
clearbutton=Button(root,image=clearImage,bd=0,bg='light blue',cursor='hand2',activebackground='light blue',command=clear)
clearbutton.place(x=590,y=550)


exitImage=PhotoImage(file='exit.png')
exitbutton=Button(root,image=exitImage,bd=0,bg='light blue',cursor='hand2',activebackground='light blue',command=iexit)
exitbutton.place(x=690,y=550)


totalLabel=Label(root,font=('new times roman',18,'bold'),bg='light blue',fg='black',)
totalLabel.place(x=6,y=560)

sentLabel=Label(root,font=('new times roman',18,'bold'),bg='light blue',fg='black',)
sentLabel.place(x=110,y=560)

leftLabel=Label(root,font=('new times roman',18,'bold'),bg='light blue',fg='black',)
leftLabel.place(x=200,y=560)

failLabel=Label(root,font=('new times roman',18,'bold'),bg='light blue',fg='black',)
failLabel.place(x=280,y=560)

























root.mainloop()