###############################################################################
##                                                                           ##
## Python code to send multiple e-mails.                                     ## 
## Written by Chammika Udalagama (chammika@nus.edu.sg)                       ## 
## ------------------------------------------------------------------------- ##
## Last updated: 2019-02-22: First version.                                  ##
##                                                                           ##
##                                                                           ##
###############################################################################

from smtplib import SMTP
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
from email.mime.base import MIMEBase

from pandas import read_excel
from datetime import datetime
from tkinter import filedialog, messagebox, Tk

global_count = 0

def sendMail(to,sender,mode='html',test = True):
    global   global_count
           
    """
    This is my file to send e-mails
    """
    
    print_message = ''
    
    #---------------------------------------------------#
    #                                                   #
    #               setup the message                   #
    #                                                   #
    #---------------------------------------------------#
    msg = MIMEMultipart()
    msg['From'] = '{} <{}>'.format(str(sender['Name']),str(sender['Email']))
    if to['Name']==None:
        msg['To'] = str(to['Email'])
    else:
        msg['To'] = '{} <{}>'.format(str(to['Name']),str(to['Email']))
        
    msg['Date'] = formatdate(localtime=True)
    msg['Subject'] = to['Subject']  #+ '({})'.format(datetime.datetime.now())
    msg.attach(MIMEText(to['Body'],mode))

    #start gathering all the e-mail addresses
    address = [to['Email']]

              
    if to['Cc'] != 'NO':
        msg['CC'] =  to['Cc']
        address += [to['Cc']]
        print_message += '\tCc: {}\n'.format(to['Cc'])
        
    if to['Bcc'] != 'NO':
        address += [to['Bcc']]
        print_message += '\tBcc: {}\n'.format(to['Bcc'])

    if to['File'] != None:
        filename = to['File'].split('/');
        filename = filename[-1]

        attachment = open(to['Path'],'rb')
        part = MIMEBase('application', 'octet-stream')
        part.set_payload((attachment).read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', "attachment; filename= %s" % filename)
        msg.attach(part)
        print_message += '\tFile: ' + str(to['Path'])
    else:
        print_message += '\tFile: None'
        
    #---------------------------------------------------#
    #                                                   #
    #              Talk to the server                   #
    #                                                   #
    #---------------------------------------------------#
    server = SMTP('smtp.nus.edu.sg',25) 
    server.ehlo()
    server.starttls()
    server.ehlo()
    server.login(sender['Userid'],sender['Password'])
    
    # If we are testing send back to the sender
    if test: 
        address = sender['Email']
        
    server.sendmail(sender['Email'],address,msg.as_string())
    server.close()
    
    print_message = '-'*30 +'\n{: 4d} Email sent \n\tto: {}\n{}\n'.format(global_count,msg['To'], print_message)
    print(print_message)
    
    # Update the log
    file = open('email-log.txt','+a')
    file.write('\n')
    if test: 
        file.write('*'*10+'TESTING'+'*'*10+'\n')
    else:
        file.write('-'*30 + '\n')
    file.write(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    file.write('\n'+ print_message)
    file.close()
        
    global_count += 1
        
def driver(info_file = 'email-info.xlsx', test = True):
    global global_count
    #------------------------------------------------------------#
    #                                                            #
    #                Get the sender's information                #
    #                                                            #
    #------------------------------------------------------------#
    sender = {'Name':None,'Email':None,'Userid':None,'Password':None}
    
    print('Reading Excel File '+ info_file)
    df = read_excel(info_file,sheet_name='Sender')
    
    for i in sender.keys():
        if i in df.columns:
            sender[i] = df.iloc[0][i]
    
    print('\tFinished reading Sender info')
    #print(sender)
    
    #------------------------------------------------------------#
    #                                                            #
    #                  Get the message template                  #
    #                                                            #
    #------------------------------------------------------------#   
    message = {'Subject':None,'Body':None,'Mode':None}
    
    df = read_excel(info_file,sheet_name='Message')
    
    for i in message.keys():
        if i in df.columns:
            message[i] = df.iloc[0][i]
            
    # Sort out the mode
    if message['Mode'] == None:
        message['Mode'] = 'text'
    
    print('\tFinished reading Message info')
    #print(message)     
    
    #------------------------------------------------------------#
    #                                                            #
    #                                                            #
    #             Send the emails to the individuals             #
    #                                                            #
    #                                                            #
    #------------------------------------------------------------#
    person_template = {'Name':None,'Email':None,'Bcc':None,'Cc':None,'Body':None, 'Subject':None,'File':None}
    
    #------------ Read in the data for the individuals ----------#
    df = read_excel(info_file,sheet_name='To')
    df.fillna(value='NO',inplace=True)
    
    print('\tFinished reading Addressee\'s info','\n')
    
    #------------------------------------------------------------#
    #              Go through, one person at a time              #
    #------------------------------------------------------------#
    global_count = 1
    for index, row in df.iterrows():
        
        # Send only three e-mails if we are testing
        if test and (index > 2):
            break
        
        person =  person_template
        
        #---------------- Get Email, Cc, Bcc and Path ---------------#
        for i in ['Email','Cc','Bcc','Path']:
            if i in df.columns and row[i] != 'None' :
                person[i] = row[i]
    
        #------------------- Construct the message ------------------#
        person['Subject'] = message['Subject']
        person['Body'] = message['Body']
        
        for r in df.columns:
            for i in ['Subject','Body']:
                person[i] = person[i].replace('{{{}}}'.format(r),str(row[r]))
        
        #---------------------- Send the email! ---------------------#
        sendMail(to = person, sender = sender, mode = message['Mode'],test=test)


if __name__=='__main__':
    root = Tk()
    root.withdraw()
    #root.update()
    filepath =  filedialog.askopenfilename(title = "Select file with E-mail Info",filetypes = (("Excel files","*.xlsx"),("all files","*.*")))
    
    #------------------- Get path to Excel file -----------------#
    if filepath:
        answer = messagebox.askyesnocancel('Send Test E-mails?',"Do you want to send a few test e-mails")
        print(answer)
        
        #--------------- Do You Want To Test Your File --------------#
        if answer == None: # Pressed 'Cancel' so quit
            exit()
        elif answer:   # Send test e-mails
            driver(info_file=filepath,test=True)
            print('Send test e-mails')
        else:
            answer = messagebox.askyesnocancel('Send All Mail for Real',"Shall I start sending all the e-mails?")
            
            #------------------- Send e-mails for real ------------------#
            if answer:
                driver(info_file=filepath,test=False)
            else:
                exit()
                
    else:
        print('No Excel maifile selected. Exitting.')
        