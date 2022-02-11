from openpyxl import *
from tkcalendar import Calendar, DateEntry
import imaplib
import email
import pywhatkit
import datetime
from tkinter import *
import requests
import json

x = datetime.datetime.now()


# Function to set focus (cursor)
def focus1(event):
	Form_gmail.focus_set()

# Function to set focus
def focus2(event):
	Form_gmailpass.focus_set()

# Function to set focus
def focus3(event):
	Form_mobileno.focus_set()


# Function to set focus
def focus4(event):
	Form_fromemail.focus_set()


# Function to set focus
def focus5(event):
	Form_Fromdate.focus_set()


# Function to set focus
def focus6(event):
	Form_todate.focus_set()


def insert():
        try:
                
	
                # if user not fill any entry
                # then print "empty input"
                if (Form_name.get() == "" and
                        Form_gmail.get() == "" and
                        Form_gmailpass.get() == "" and
                        Form_mobileno.get() == "" and
                        Form_fromemail.get() == "" and
                        Form_Fromdate.get() == "" and
                        Form_todate.get() == ""):
                                
                        print("empty input")

                else:
                        #get input values                         
                        frm1_username = Form_name.get()
                        frm1_mailid = Form_gmail.get()
                        frm1_gmailpass = Form_gmailpass.get()
                        frm1_mobilnumber = Form_mobileno.get()
                        frml_fromemail = Form_fromemail.get()
                        frm1_fromdate = Form_Fromdate.get()
                        frml1_todate = Form_todate.get()
                        

                        try:
                                def listToString(s):
                                    # initialize an empty string
                                    str1 = "\n"
            
                                    # return string
                                    return (str1.join(s))
                                
                                print(x)
                                hourtiming=int(x.strftime("%H"))
                                print(hourtiming)
                                minutetiming=int(x.strftime("%M"))+2

                                #credentials
                                username =frm1_mailid
                                #generated app password
                                app_password= frm1_gmailpass
                                gmail_host= 'imap.gmail.com'
                                #set connection
                                try:
                                        #mail connections
                                        mail = imaplib.IMAP4_SSL(gmail_host)
                                        #login
                                        mail.login(username, app_password)
                                        textwhat=[]
                                        textsms=[]                                        
                                        textwhat.append("Dear "+frm1_username+","+"\n")
                                        #select inbox
                                        mail.select("INBOX")



                                        
                                        #select specific mails
                                        _, selected_mails = mail.search(None, '(FROM ' + frml_fromemail +')')
                                        #total number of mails from specific user
                                        print("Total Messages "+username+":" , len(selected_mails[0].split()))
                                        for num in selected_mails[0].split():
                                                _, data = mail.fetch(num , '(RFC822)')
                                                _, bytes_data = data[0]

                                                #convert the byte data to message
                                                email_message = email.message_from_bytes(bytes_data)
                                                print("\n========================")
                                                #access data
                                                print("Subject: ",email_message["subject"])
                                                print("To:", email_message["to"])
                                                print("From: ",email_message["from"])
                                                print("Date: ",email_message["date"])
                                                for part in email_message.walk():
                                                    if part.get_content_type()=="text/plain" or part.get_content_type()=="text/html":
                                                        message = part.get_payload(decode=True)
                                                        print("Message: \n", message.decode())
                                                        moh=message.decode()
                                                        print("=======================\n")

                                                        textwhat.append("\n========================\n"+"Subject: "+email_message["subject"]+"\n" + "To        :"+email_message["to"] +"\n"+ "From    : "+email_message["from"]+"\n" + "Date     : "+email_message["date"]+"\n message :\n\t"+moh+"\n=====================\n")
                                                        textsms.append("Subject: "+email_message["subject"] + "Date     : "+email_message["date"]+"\n message :\n\t"+moh)


                                                        break
                                except imaplib:
                                        print("check connections")

                                #sending whatsapp message
                                textwhatmsg=listToString(textwhat)
                                pywhatkit.sendwhatmsg("+91"+frm1_mobilnumber, textwhatmsg, hourtiming,minutetiming)

                                textsmsmsg=listToString(textsms)                               

                                #Sending Normal Message
                                try:
                                        url = "https://www.fast2sms.com/dev/bulk"
                                        # create a dictionary 
                                        my_data = { 
                                            # Your default Sender ID 
                                            'sender_id': 'FSTSMS', 
                                            
                                            # Put your message here! 
                                            'message': textsmsmsg, 
                                            
                                            'language': 'english', 
                                            'route': 'p', 
                                            
                                            # You can send sms to multiple numbers 
                                            # separated by comma. 
                                            'numbers':  frm1_mobilnumber  
                                        } 

                                        # create a dictionary 
                                        headers = { 
                                            'authorization': 'zRQsrjOhwkmKIJcxbFPuHLSW6tqVC8T3BXv9p5MyoD7eEYad1gzwOlEJfBybZuphYT5a4s02KHdxQkMi', 
                                            'Content-Type': "application/x-www-form-urlencoded", 
                                            'Cache-Control': "no-cache"
                                        }

                                        # make a post request 
                                        response = requests.request("POST",url,data = my_data,headers = headers) 

                                        returned_msg = json.loads(response.text) 

                                        # print the send message 
                                        print(returned_msg['message'])
                                except:
                                        print("error on message sending")
                        except:
                                print("Problem on mail connection to the server")    
        except:
                print("Error")
# Driver code
if __name__ == "__main__":
        # create a GUI window
	root = Tk()

	# set the background colour of GUI window
	root.configure(background='light green')

	# set the title of GUI window
	root.title("registration form")

	# set the configuration of GUI window
	root.geometry("500x300")


	# create a Form label
	heading = Label(root, text="Gmail Filter and Message sending script", bg="light green")

	# create a Name label
	Frm_name = Label(root, text="Name", bg="light green")
	frm_gmail = Label(root, text="G-Mail", bg="light green")
	frm_gmailpass = Label(root, text="G-Mail Pass", bg="light green")
	frm_mobileno= Label(root, text="Mobile No", bg="light green")
	frm_fromemail = Label(root, text="From Email", bg="light green")
	frm_fromdate = Label(root, text= "From Date",  bg="light green")
	frm_todate = Label(root, text= "To Date",  bg="light green")

	# grid method is used for placing
	# the widgets at respective positions
	# in table like structure .
	heading.grid(row=0, column=1)
	Frm_name.grid(row=1, column=0)
	frm_gmail.grid(row=2, column=0)
	frm_gmailpass.grid(row=3, column=0)
	frm_mobileno.grid(row=4, column=0)
	frm_fromemail.grid(row=5, column=0)
	frm_fromdate.grid(row=6, column=0)
	frm_todate.grid(row=7, column=0)

	# create a text entry box
	# for typing the information
	Form_name = Entry(root)
	Form_gmail = Entry(root)
	Form_gmailpass = Entry(root)
	Form_mobileno = Entry(root)
	Form_fromemail = Entry(root)
	Form_Fromdate = DateEntry(root, width= 16, background= "magenta3", foreground= "white",bd=2)
	Form_todate = DateEntry(root, width= 16, background= "magenta3", foreground= "white",bd=2)
	# bind method of widget is used for
	# the binding the function with the events

	Form_name.bind("<Return>", focus1)
	Form_gmail.bind("<Return>", focus2)	
	Form_gmailpass.bind("<Return>", focus3)	
	Form_mobileno.bind("<Return>", focus4)
	Form_fromemail.bind("<Return>", focus5)
	Form_Fromdate.bind("<Return>", focus6)

	# grid method is used for placing
	# the widgets at respective positions
	# in table like structure .
	Form_name.grid(row=1, column=1, ipadx="100")
	Form_gmail.grid(row=2, column=1, ipadx="100")
	Form_gmailpass.grid(row=3, column=1, ipadx="100")
	Form_mobileno.grid(row=4, column=1, ipadx="100")
	Form_fromemail.grid(row=5, column=1, ipadx="100")
	Form_Fromdate.grid(row=6, column=1, ipadx="100")
	Form_todate.grid(row=7, column=1, ipadx="100")


	# create a Submit Button and place into the root window
	submit = Button(root, text="Submit", fg="Black",bg="Red", command=insert)
	submit.grid(row=8, column=1)

	# start the GUI
	root.mainloop()
