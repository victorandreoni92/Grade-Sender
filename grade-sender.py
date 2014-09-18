#!/usr/bin/env python
from smtplib import SMTP
from smtplib import SMTPException
from email.MIMEMultipart import MIMEMultipart
from email.MIMEText import MIMEText
from email.MIMEImage import MIMEImage
from email.MIMEBase import MIMEBase
from email import Encoders
import sys
import fnmatch
import os
import time

"""
Created by Victor Andreoni
victorandreoni92 at gmail com
"""


# Global variables 

PROGRESS_BAR_WIDTH = 12
GRADES_DIRECTORY = "grades/"

EMAIL_FROM = "wpics2303a14@gmail.com" # The address ued to send the emails
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
TEXT_SUBTYPE = "plain"

GLOBAL_WIDTH = 80

def build_email_body(grader_name, assignment_name):
    """
    Builds the content of the email body based on the parameters given
    """
    content = "Hello.\n\nHere is your grade for %1s.\n\n%2s" % (assignment_name, grader_name)
    return content

def build_email(student_email, grader_email, grader_name, assignment_name, filename):
    """
    Builds an email with the information provided. The email is then sent by send_email()
    """
    # Set up the email recipients
    recipients = []
    recipients.append(student_email)
    recipients.append(grader_email)
 
    # Build the email
    message = MIMEMultipart()
    message["Subject"] = assignment_name + " Grade"
    message["From"] = EMAIL_FROM
    message["To"] = ", ".join(recipients)
    message.add_header('reply-to', grader_email)
    
    message_body = MIMEMultipart('alternative')
    message_body.attach(MIMEText(build_email_body(grader_name, assignment_name), TEXT_SUBTYPE ))
    
    attachment = MIMEBase('application', 'zip')
    attachment.set_payload(open(GRADES_DIRECTORY + filename).read())
    Encoders.encode_base64(attachment)
    attachment.add_header('Content-Disposition', 'attachment', filename=filename)

    # Attach the message body
    message.attach(message_body)
    # Attach grades
    message.attach(attachment) 

    # Return the message and the list of recipients
    return (message, recipients)

def send_email(pswd, message, recipients):
    """
    Sends the email previously built
    """  
    try:
      smtpObj = SMTP(SMTP_SERVER, SMTP_PORT)
      # Identify to server
      smtpObj.ehlo()
      # Switch SMTP connection to TLS mode and identify again
      smtpObj.starttls()
      smtpObj.ehlo()
      # Login
      smtpObj.login(user=EMAIL_FROM, password=pswd)
      # Send email
      smtpObj.sendmail(EMAIL_FROM, recipients, message.as_string())
      # Close connection and session.
      smtpObj.quit()
    except SMTPException, error:
      print "Error: could not send email"
 
def buildProgressBar():
    """
    Sets up the progress bar by printing the edges to stdout and setting the cursor
    to the start of the bar
    """	    
    sys.stdout.write("| [%s]%s|" % (" " * PROGRESS_BAR_WIDTH, " " * (GLOBAL_WIDTH - PROGRESS_BAR_WIDTH - 3)))
    sys.stdout.flush()
    # Return to start of line
    sys.stdout.write("\b" * (GLOBAL_WIDTH - 1))

def updateOneBarTick():
    """
    Updates the progress bar by one step.
    """
    sys.stdout.write("*")
    sys.stdout.flush()

def endProgressBar():
    """
    Ends the progress bar by printing a new line
    """
    sys.stdout.write("\n")
    sys.stdout.flush()

def displayScriptInfo():
    """
    Displays the initial script information
    """
    sys.stdout.write("|%s|" % ("-" * GLOBAL_WIDTH))
    sys.stdout.write("\n")
    
    welcome_message = " Welcome to Grade Sender 1.0"
    sys.stdout.write("|%s%s|" % (welcome_message, " " * (GLOBAL_WIDTH - len(welcome_message))))
    sys.stdout.write("\n")
    sys.stdout.write("|%s|" % (" " * GLOBAL_WIDTH))
    sys.stdout.write("\n")

    info_message = " The script will now send the files located in the grades directory and"
    info_message2 = " will update you on its progress." 
    sys.stdout.write("|%s%s|" % (info_message, " " * (GLOBAL_WIDTH - len(info_message))))
    sys.stdout.write("\n")
   
    sys.stdout.write("|%s%s|" % (info_message2, " " * (GLOBAL_WIDTH - len(info_message2))))
    sys.stdout.write("\n")

    sys.stdout.write("|%s|" % ("-" * GLOBAL_WIDTH))
    sys.stdout.write("\n")
   
    sys.stdout.flush()

def displayScriptEnd():
    """
    Notifies the user that all grades were sent
    """
    sys.stdout.write("|%s|" % ("-" * GLOBAL_WIDTH))
    sys.stdout.write("\n")
    
    end_message = " All grades have been sent!"
    sys.stdout.write("|%s%s|" % (end_message, " " * (GLOBAL_WIDTH - len(end_message))))
    sys.stdout.write("\n")
    
    sys.stdout.write("|%s|" % ("-" * GLOBAL_WIDTH))
    sys.stdout.write("\n")

def main(pswd, grader_email, grader_name):
    """
    This function loops through all the files in the grades directory and emails them to the user,
    based on the filename. All users are assumed to have a @wpi email address
    """
    # Initialize counters
    count = 0
    totalFiles = 0

    displayScriptInfo()

    # Loop through all files in the grades directory
    for dirname, dirnames, filenames in os.walk(GRADES_DIRECTORY):
	totalFiles = len(filenames)
       
        # Loop only through the files, ignore directories
        for filename in filenames:
            count += 1

            # Display progress
            send_message = " Sending email %d of %d" % (count, totalFiles)
            sys.stdout.write("|%s%s|" % (send_message, " " * (GLOBAL_WIDTH - len(send_message))))
            sys.stdout.write("\n")

            # Signal begining of processing
            buildProgressBar()
	    updateOneBarTick()
            
            # Prepare user email address and call build_email function
            filename_components = filename.split('-', 2)

            # Set the assignment name and student email address
            assignment_name = filename_components[1]           
            student_email = filename_components[0] + "@wpi.edu"
	    
            # Build the email and sent it
            message, recipients = build_email(student_email, grader_email, grader_name, assignment_name, filename)
            send_email(pswd, message, recipients)
            
            # After the email was sent, add a tick to the progress bar
            updateOneBarTick()
            
            # Wait some time to avoid hammering the server
	    for i in xrange(10,0,-1):
                time.sleep(1)
                updateOneBarTick()
            
            endProgressBar()
        
        displayScriptEnd()

    if (count == 0):
        print "\nNo files were found. Is the %s folder on the same directory as the script?\n" % (GRADES_DIRECTORY)

def usage():
    print  """
******************************************************************************************
* Correct usage: python grade-sender.py 'senderEmailPassword' 'graderEmail' 'graderName' *
*     senderEmailPassword: The password for the account being used to send the emails    *
*     graderEmail: The grader's email used as reply-to address                           *
*     graderName: The grader's name used as a signature on the email                     *
******************************************************************************************
           """


# Script called from command-line, call main function.
if __name__ == "__main__":
    if len(sys.argv) == 4:
        main(sys.argv[1], sys.argv[2], sys.argv[3]);
    else:
        usage()
    
