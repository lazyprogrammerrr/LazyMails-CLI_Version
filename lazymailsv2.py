#imports
import pandas as p
import smtplib,ssl
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.message import EmailMessage
import docx2txt
import os
from rich.console import Console
from rich.table import Table
from rich import print
import time
import stdiomask
from email import encoders
import getpass

c = Console() # instance of rich module 
msg = EmailMessage()

logo = '''

                            ██╗░░░░░░█████╗░███████╗██╗░░░██╗███╗░░░███╗░█████╗░██╗██╗░░░░░░██████╗
                            ██║░░░░░██╔══██╗╚════██║╚██╗░██╔╝████╗░████║██╔══██╗██║██║░░░░░██╔════╝
                            ██║░░░░░███████║░░███╔═╝░╚████╔╝░██╔████╔██║███████║██║██║░░░░░╚█████╗░
                            ██║░░░░░██╔══██║██╔══╝░░░░╚██╔╝░░██║╚██╔╝██║██╔══██║██║██║░░░░░░╚═══██╗
                            ███████╗██║░░██║███████╗░░░██║░░░██║░╚═╝░██║██║░░██║██║███████╗██████╔╝
                            ╚══════╝╚═╝░░╚═╝╚══════╝░░░╚═╝░░░╚═╝░░░░░╚═╝╚═╝░░╚═╝╚═╝╚══════╝╚═════╝░
'''
print(logo)
print("\t\t\t\t\t\t  Git - LazyProgrammerrr")
print("\nYou Can Press [b bright_red] (ctrl + c) [/b bright_red] to exit anytime.")

#try block with loop to interpret data from excel and setting up server
try:
    while True:
        email_input = c.input("\n[b magenta]Enter Your Email Here -> [/b magenta]")
        #password_input = c.input("\n[b bright_yellow]Enter Your Password Here -> [/b bright_yellow]")
        password_input = stdiomask.getpass(prompt='\nEnter Your Password Here -> ')
        context = ssl.create_default_context()
        port = 465
        smtp_server = "smtp.gmail.com"
        with smtplib.SMTP_SSL(smtp_server,port,context=context) as server:
            try:
                server.login(email_input,password_input)
                print("\n[b bright_green]Logged In![/b bright_green]")
                break
            except:
                print("\n[b bright_red]Wrong Email Id or Password! Please Check Again.[/b bright_red]\n")

    print("\n(Email List) Choose from XLSX or CSV file!")
    while True:
        file_choose = c.input("\nEnter Here [b bright_red](xlsx) or (csv)[/b bright_red]-> ").casefold()
        if file_choose == 'xlsx':
            file_input = c.input("\n[b bright_cyan]Choose Your Email List [b bright_yellow](XLSX)[/b bright_yellow]File Here \nYou Can Also Paste The Path Link Of File [b bright_red]'Using Shift + Right Click'[/b bright_red]-> [/b bright_cyan]").replace('"', '')
            try:
                datafile = p.read_excel(file_input)
                contacts = p.DataFrame(datafile)
                break
            except Exception as e:
                print("[b bright_red]File Not Found! or Selected File Is Not of (XLSX) Format![/b bright_red]")
        if file_choose == 'csv':
            file_input = c.input("\n[b bright_cyan]Choose Your Email List [b bright_yellow](CSV)[/b bright_yellow]File Here \nYou Can Also Paste The Path Link Of File [b bright_red]'Using Shift + Right Click'[/b bright_red]-> [/b bright_cyan]").replace('"', '')
            try:
                datafile = p.read_csv(file_input)
                contacts = p.DataFrame(datafile,columns = ['Name','Email'])
                break
            except Exception as e:
                print("[b bright_red]File Not Found!  or Selected File Is Not of (CSV) Format![/b bright_red]")
        else:
            print("\nPlease Choose From Given Options Only!")

    
    subject_input = c.input("\n[b bright_blue]Enter Your Subject Here-> [/b bright_blue]")
    while True:
            template_input = c.input("\n[b grey100]Enter Your Body Template [b bright_yellow](DOCX)[/b bright_yellow]File Name Here \nYou Can Also Paste The Path Link Of File [b bright_red]'Using Shift + Right Click'[/b bright_red]  -> [/b grey100]").replace('"', '')
            try:
                template = docx2txt.process(template_input)
                break
            except:
                print("[b bright_red]File Not Found![/b bright_red]")
    try:
        attachement_ask = c.input("\n[b bright_yellow]Do You Want To Add Attachement To Your Email? Press(Y) or (N) -> [/b bright_yellow]").casefold()
        if attachement_ask == 'y':
            while True:
                attach_file_name = c.input("\n[b magenta]Enter Attachement File Path Here -> [/b magenta]").replace('"', '')
                attachfilenameask = c.input("\n[b blue]Enter Custom File Name Here [b green](WITH FILE EXTENSTION)[/b green] -> [/b blue]")
                try:
                    attach_file = open(attach_file_name, 'rb') # Open the file as binary mode
                    file_name = attach_file.name
                    payload = MIMEBase('application', 'octet-stream',Name=attachfilenameask)
                    payload.set_payload((attach_file).read())
                    encoders.encode_base64(payload) #encode the attachment
                    payload.add_header('Content-Disposition', 'attachment')
                    print("[b bright_green]Attachement Added![/b bright_green]\n")
                    yes = True
                    # message.attach(payload)
                    time.sleep(0.2)
                    print("\nSending Mails:")
                    break
                except:
                    print("[b bright_red]File Not Found![/b bright_red]")
        else:
            yes = False
            print("[b bright_red]No Attachement Is Added![/b bright_red]")
            time.sleep(0.2)
            print("\nSending Mails:")
    except:
        print("Some Error Occured")
    
    for i in range(len(contacts)):
        Name , Email = contacts.iloc[i] #unpacking the data
        message = MIMEMultipart("alternative")
        message['From'] = email_input
        message['To'] = Email
        message['Subject'] = subject_input
        message.attach(MIMEText(template.format(Name.split()[0]),'html'))
        
        if yes == True:
            message.attach(payload)
        else:
            pass

        text = message.as_string()
        context = ssl.create_default_context()

        with smtplib.SMTP_SSL(smtp_server,port,context=context) as server:
            server.login(email_input,password_input)
            server.sendmail(message["From"],message["To"],text)
            server.quit()
        time.sleep(0.1)
        path = os.path.join('C:\\','Users',getpass.getuser(),'Documents','LazyMailsLog')
        if not os.path.exists(path):
            os.makedirs(path)
        filename = 'Lazymaillogs.txt'
        with open(os.path.join(path, filename), 'a') as f:
                x = datetime.now()
                f.write("Log Time-> "+ str(x) + "\tName-> " + Name + " Email -> " + Email + "\n")
        print("\n[b bright_cyan]Sent To -> [/b bright_cyan]",Name,'[b bright_yellow]Id: [/b bright_yellow]'+Email)
    with open(os.path.join(path, filename), 'a') as f:
        f.write("\n\n\t\t\t------ END OF LOG ------\n\n")
    print("\n[b bright_green]All Email Sent[/b bright_green]")
    time.sleep(2)

#handling exception
except KeyboardInterrupt:
    print("\n\n[b bright_red]Program Exiting.[/b bright_red]")
    time.sleep(1)