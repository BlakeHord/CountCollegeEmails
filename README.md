# CountCollegeEmails
Counting the number of emails I receive from colleges each day

## How it works:
* Uses the python imaplib and email libraries to read a gmail account
* Then uses the openpyxl python library to modify an excel document and add a new line to it
* A second excel document is set up to read the data from the first and automatically make graphs from it when opened

## Get set up
You will need to make a file named password.py with the following format
'''
username = "put your username here"
password = "put your password here"
'''
Just those two lines and it will read it. I am working on making it more secure but this does the trick.

## Files this repository contains
* CountSchedule.py
  * executes the search and logging every night at 3:27 AM, a time which can be changed to fit your needs.
* gmailconnect.py
  * the basics of reading and counting from a gmail inbox. The subroutines used in this scirpt were transferred over to CountSchedule.py
* CollectSpamEmail.py
  * the same as gmailconnect.py but with integration with the excel document, but without the scheduling of CountSchedule.py

## Parameters of the search
* The script searched the inbox named "All Mail" for every email received in the day before it is executed. 
* It searches at 3:27 AM or whenever the computer is on next
* It looks for the word "admission" in the subject, sender's email, or body of the email and if it does the count is increased by one
* I didn't want emails from Khan Academy included in the count, because I only wanted the emails from colleges included. So, I added a "blacklist" of terms used in the sender's email that would disqualify the email from being counted. 

If you have any suggestions on how to improve the quality of the count, including which search terms I should use, please let me know.