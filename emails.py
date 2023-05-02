from email.message import EmailMessage
import smtplib
import os

def mailing_list():
    """Updates the mailing list for each email group. The email groups 
    include section groups, role groups, and all singers."""
    # Not sure about this function, might not use it.

def send_email(reciever_email, subject, body):
    """Sends an email to a member in the spreadsheet. Planning to use for 
    update confirmations, semesterly prompts for updates, 
    new member welcome email"""
    
    sender_email = "vopmembershiptest@gmail.com"
    password = os.environ.get("VOP_PASSWORD")
    
    #Create an instance of the EmailMessage class to generate the email
    em = EmailMessage()
    em['From'] = sender_email
    em['To'] = reciever_email
    em['Subject'] = subject
    em.set_content(body)
    
    # Log in to sender_email and define SMTP server
    smtp_server = 'smtp.gmail.com'
    smtp_port = 587 # SMTP over TLS
    
    # Connect to email server using TLS
    server = smtplib.SMTP(smtp_server, smtp_port)
    server.ehlo()
    server.starttls()
    server.login(sender_email, password)

    # Send the email
    server.sendmail(sender_email, reciever_email, em.as_string())
    server.quit()

def new_member_email_body(file, receiver_email):
    """Fills in an existing template for a new member email with the new 
    member's name, section, and section leader, creating a new file with
    the body of the email and returning the contents of that file as a str."""
    pass

