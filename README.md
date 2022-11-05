Hi folks Tejas Jadhav here and this is my first project.
This program can be helpful in editing and sending the participation certificates when we have large number of participants.
My program does the following:
1. Extract name and email ids of participants from the excel sheet. The excel sheet should be in a particular format refer the certificate.xlsx file
2. Later replaces the ### in the sample.docx file with name of participants.
3. Sends the certificates to the respective email ids of participants.

To implement this program in your machine you need to make the following changes:
1. Import all the jar files present in the lib folder.
2. Change the sample.docx file according to your needs just add '###' where the name of participants needs to be entered.
3. Make changes in the excel file . But 1st column for name 2nd for surname and the 3rd column for email id and no spaces between rows.

Changes in code are as follows:
1. Line 30:  Change the filepath.
2. Line 78: Change the filepath.
3. Line 86: Your email id from wich you need to send the email.
4. Line 87: password of your email id.
5. Line 88: your email id from which you need to send the email.
6. Line 110: Subject of mail.
7. Line 116: Body of the mail
8. Line 123: Change the filepath.

This program needs strong internet connection for sending the mails.
