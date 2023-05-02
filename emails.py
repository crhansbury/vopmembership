from email.message import EmailMessage
import smtplib
import os
from querymember import query_member_attr

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
    print("Email successfully sent to {}".format(reciever_email))

def new_member_email(template, receiver_email, spreadsheet):
    """Returns two values: (subject, email_body). Fills in an existing template 
    for a new member email with the new 
    member's name, section, and section leader, creating a new file with
    the body of the email and returning the contents of that file as a str.
    Also returns an appropriate subject line for the email."""
    # Retrieving the attributes from the spreadsheet for the new member
    new_mem_attributes = query_member_attr(spreadsheet, "email", receiver_email, \
                                           "first_name", "last_name", "pronouns", \
                                            "section", "id", "phone", "address", \
                                            "city", "state", "zip")
    nma_str = [str(i) for i in new_mem_attributes[0]]
    # Retrieving information about all section leaders from the spreadshet
    sec_leader_attributes = query_member_attr(spreadsheet, "role", "Section leader", \
                                            "section", "first_name", "last_name")
    # Finding the section leader for the new member's section
    
    section_leader = []
    for list in sec_leader_attributes:
        new_section = nma_str[3]
        if list[0].startswith(new_section[0]):
            section_leader.extend(list)
    # Open the email template and replace all variables with the correct info
    subject = "Welcome to VOP, {}!".format(nma_str[0])
    with open(template) as f:
        email_template = f.read()
    email_body = email_template.format(first_name=nma_str[0], \
                                        last_name=nma_str[1], \
                                        pronouns=nma_str[2], \
                                        section=nma_str[3], \
                                        id=nma_str[4], \
                                        email=receiver_email, \
                                        phone=nma_str[5], \
                                        address=nma_str[6], \
                                        city=nma_str[7], \
                                        state=nma_str[8], \
                                        zipcode=nma_str[9], \
                                        sec_first=section_leader[1], \
                                        sec_last=section_leader[2])
    return subject, email_body