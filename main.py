'''
Author: Defrag_Defector
Purpose: From a CSV file take the name of a client, search and replace name in file template
         then create a outlook instance and insert desired values, send the email.
'''
import win32com
import win32com.client 
import mammoth
import os
import codecs
import csv
import time

def make_doc(name): #takes docx file with email template, converts it to HTML, then target word in email is found and then replaced
    with open("MessageBody.docx", "rb") as docx_file: 
            result = mammoth.convert_to_html(docx_file) 
            html_ver = result.value
            with open("tmp.html","w") as converted: #take converted docx and paste it to tmp file
                converted.write(html_ver)
            with open('tmp.html',mode='r') as replacer: #find attribute to replace '[name]'
                data = replacer.read()
                data = data.replace("[name]", name)
            with open('tmp.html','w') as replaced: #replace found instnace of [name] with name
                replaced.write(data)
    return data
      
with open('test.csv','r') as contacts_file:
    name_reader = csv.reader(contacts_file, delimiter=",")
    for row in name_reader:
        name_of_contact = row[0]     
        body_to_send = make_doc(name_of_contact)
        #Emial signature code: Ripped from https://stackoverflow.com/questions/32209091/add-signature-to-outlook-email-with-python-using-win32com
        sig_files_path = 'AppData\Roaming\Microsoft\Signatures\\signature_files\\' #Default values yours may vary
        sig_html_path = 'AppData\Roaming\Microsoft\Signatures\\signature.htm'
        signature_path = os.path.join((os.environ['USERPROFILE']), sig_files_path) # Finds the path to Outlook signature files with signature name "Work"
        html_doc = os.path.join((os.environ['USERPROFILE']),sig_html_path) #Specifies the name of the HTML version of the stored signature
        html_doc = html_doc.replace('\\\\', '\\')
        html_file = codecs.open(html_doc, 'r', 'utf-8', errors='ignore') #Opens HTML file and converts to UTF-8, ignoring errors
        signature_code = html_file.read()  #Writes contents of HTML signature file to a string
        signature_code = signature_code.replace(('signature_files/'), signature_path) #Replaces local directory with full directory path
        html_file.close()
        #Email Building and Sending Logic
        outlook = win32com.client.Dispatch("Outlook.Application")
        namspace = outlook.GetNamespace("MAPI")
        mail_Item = outlook.CreateItem(0)
        mail_Item.Subject = "subject Text"
        mail_Item.BodyFormat = 2 #1 is plain, 2 is HTML, 3 is rich
        mail_Item.HTMLBody = body_to_send + signature_code #Pastes HTML of body and signature to outlook instance
        mail_Item.GetInspector
        mail_Item.To = row[1] # email of client
        mail_Item.CC = 'ccEmailExample@example.com'
        mail_Item.Display() #Shows draft breifly before being sent NEEDS TO BE ON TO ALLOW THE DRAFT TO BE SENT
        #mail_Item.Send() MAKE SURE EVERYTHING IS RIGHT BEFORE STARTING
        print("Message Sent to: " + row[0] + " at " + row[1] + " Now Sleeping for 15 minutes")
        time.sleep(900)
        
