#imports
import pandas as p
import smtplib,ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import docx2txt
import os
from rich.console import Console
from rich.table import Table
from rich import print
import time
import stdiomask

c = Console() # instance of rich module 

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

#try block with loop to interpret data from excel and setting up server
try:
    while True:
        email_input = c.input("[b magenta]Enter Your Email Here -> [/b magenta]")
        # password_input = c.input("\n[b bright_yellow]Enter Your Password Here -> [/b bright_yellow]")
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
    while True:
        file_input = c.input("\n[b bright_cyan]Choose Your Email List File Here \nYou Can Also Paste The Path Link Of File [b bright_red]'Using Shift + Right Click'[/b bright_red]-> [/b bright_cyan]").replace('"', '')
        try:
            datafile = p.read_excel(file_input)
            contacts = p.DataFrame(datafile)
            break
        except:
            print("[b bright_red]File Not Found![/b bright_red]")
            
    subject_input = c.input("\n[b bright_blue]Enter Your Subject Here-> [/b bright_blue]")
    while True:
            template_input = c.input("\n[b grey100]Enter Your Body Template File Name Here \nYou Can Also Paste The Path Link Of File [b bright_red]'Using Shift + Right Click'[/b bright_red]  -> [/b grey100]").replace('"', '')
            try:
                template = docx2txt.process(template_input)
                break
            except:
                print("[b bright_red]File Not Found![/b bright_red]")
    for i in range(len(contacts)):
        Name , Email = contacts.iloc[i] #unpacking the data
        message = MIMEMultipart("alternative")
        message['From'] = email_input
        message['To'] = Email
        message['Subject'] = subject_input
        message.attach(MIMEText(template.format(Name.split()[0]),'html'))
        text = message.as_string()
        context = ssl.create_default_context()

        with smtplib.SMTP_SSL(smtp_server,port,context=context) as server:
            server.login(email_input,password_input)
            server.sendmail(message["From"],message["To"],text)
            server.quit()
        time.sleep(0.1)
        print("[b bright_cyan]Sent To -> [/b bright_cyan]",Name,'[b bright_yellow]Id: [/b bright_yellow]'+Email)
    print("\n[b bright_green]All Email Sent[/b bright_green]")
    time.sleep(2)

#handling exception
except KeyboardInterrupt:
    print("\n\n[b bright_red]You Pressed Wrong Keys! Program Exiting.[/b bright_red]")
    time.sleep(1)