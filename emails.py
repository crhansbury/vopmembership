from email.message import EmailMessage
import smtplib
import os
from querymember import query_member_attr

def mailing_list():
    """Updates the mailing list for each email group. The email groups 
    include section groups, role groups, and all singers."""
    # Not sure about this function, might not use it.

def send_email(reciever_email, subject, body):
    """Sends an email to the email defined in the 'reciever_email' parameter.
    Sends from vopmembershiptest@gmail.com. Password is defined in a local
    environment variable and cannot be retrieved outside of the test environment
    at this time."""
    
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
    
    # Connect to email server using TLS and log in to email
    server = smtplib.SMTP(smtp_server, smtp_port)
    server.ehlo()
    server.starttls()
    server.login(sender_email, password)

    # Send the email
    server.sendmail(sender_email, reciever_email, em.as_string())
    server.quit()
    print("Email successfully sent to {}".format(reciever_email))

def generate_email(template, receiver_email, spreadsheet):
    """Returns two values: (subject, email_body). Fills in an existing template 
    for an email with all of the memeber's attributes, their section leader, and
    the section leader's email. 
    
    The subject is defined in the template as the first line followed by an 
    empty line. The body includes the third line of the template to the end. 
    
    The variables in the template include all attributes of the member 
    associated with the 'receiver_email' parameter passed into the function as 
    defined in the spreadsheet file passed into the 'spreadsheet' parameter. 
    The variables must be written in the template as: {first_name}, {last_name}, 
    {pronouns}, {section}, {id}, {email}, {phone}, {address}, {city}, {state}, 
    {zipcode}. Other variables inclue the section, name, and email of the 
    section leader associated with the member's section. These variables must
    be written in the template as: {sec_first}, {sec_last}, and {sec_email}."""
    
    # Retrieving the attributes from the spreadsheet for the member
    member_attributes = query_member_attr(spreadsheet, "email", receiver_email,
                                           "first_name", "last_name", "pronouns",
                                            "section", "id", "phone", "address",
                                            "city", "state", "zip")
    # Turning the items in the list of attributes into strings
    mattr_string = [str(i) for i in member_attributes[0]]
    # Retrieving attributes for all section leaders from the spreadshet
    sec_leader_attributes = query_member_attr(spreadsheet, "role",
                                              "Section leader", "section",
                                              "first_name", "last_name", "email")
    # Finding the section leader for the member's section
    section_leader = []
    for list in sec_leader_attributes:
        new_section = mattr_string[3]
        if list[0].startswith(new_section[0]):
            section_leader.extend(list)
    # Open the email template and replace all variables with the correct info
    with open(template) as f:
        subject = f.readline().strip().format(first_name=mattr_string[0])
        next(f)
        email_template = f.read()
    email_body = email_template.format(first_name=mattr_string[0],
                                        last_name=mattr_string[1],
                                        pronouns=mattr_string[2],
                                        section=mattr_string[3],
                                        id=mattr_string[4],
                                        email=receiver_email,
                                        phone=mattr_string[5],
                                        address=mattr_string[6],
                                        city=mattr_string[7],
                                        state=mattr_string[8],
                                        zipcode=mattr_string[9],
                                        sec_first=section_leader[1],
                                        sec_last=section_leader[2],
                                        sec_email=section_leader[3])
    return subject, email_body