# outlook-based-email-bot

## Goal

The goal was to utilize python to automate sending emails via outlook. 

## Modules Used

###Mammoth

for converting Docx files to HTML.

###win32com

for opening and sending emails through outlook.

###csv

two coulmn CSV files was used to store names and the respective email address for that contact. 

###time

Setting a 15 minute timmer to avoid suspicion from the outlook domain owner.

## Limitations

One thing I had to do was copy some code from slackoverflow to add my signature to each email. 
The code was modified slightly to aid my efforts. Second, you have to be on a machine that has Outlook installed
on the machine. Because the account holder is my University, I didn't want to bother asking for an auth token if it was
needed.

## Other Note

I am not a Comp. Sci. student. But a security student so if my code sucks I'd love to know how I could make it better for the future.
