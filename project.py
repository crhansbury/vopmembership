from openpyxl import load_workbook
from datetime import datetime
from emails import generate_email, send_email
from newmember import create_member_id
from querymember import query_active_member_attr

def add_member(file):
    """Adds a member to the spreadsheet. Takes criteria for:
    name, pronouns, voice part, role, email address, phone number, and 
    mailing address. Appends spreadsheet with this information. Updates
    a field that indicates the date of the latest update."""
    # Prompt the user for info
    first_name = input("First Name: ")
    last_name = input("Last Name: ")
    pronouns = input("Pronouns: ")
    section = input("Section: ")
    role = input("Role: ")
    email = input("Email: ")
    phone = input("Phone Number: ")
    address = input("Street Address: ")
    city = input("City: ")
    state = input("State: ")
    zipcode = input("Zip Code: ")
    # Generate new member ID
    try:
        member_id = create_member_id(file)
    except:
        print("{} not found, or is an invalid format. Please enter a valid filename.".format(file))
        exit()
    # Open the workbook and load the "Active Member" sheet and append info
    workbook = load_workbook(filename=file)
    sheet = workbook["Active Members"]
    rows = (
        last_name,first_name,pronouns,section,member_id,role,address,city, \
        state,zipcode,phone,email
    )
    sheet.append(rows)
    # Adding the date modified
    date = datetime.now()
    sheet["M1"] = "Date Modified: {} at {}".format(date.strftime("%m/%d/%Y"), \
                                                date.strftime("%H:%M"))
    workbook.save(file)
    print("{} {} successfully added to {}.".format(first_name, last_name, file))
    # Generating and sending welcome email once new member has been successfully
    # appended to spreadsheet
    try:
        subject, body = generate_email("email-templates/new_member_template.txt", email, file)
        send_email(email, subject, body)
    except:
        print("Email to {} was unsuccessful.".format(email))

def mailinglist_email(template, spreadsheet, attribute, attr_value):
    """Ideally: sends an email to a certain specified group of memebrs
    in the active members sheet. could send to all emails, just section
    leaders, just people with roles, or just people in a certain section."""
# Ok I cannot figure out how to be able to send emails to everyone AND to only
# a section or group. This is possbily beyond the scope of the project right now.
# Might come back to it later.

    active_members = query_active_member_attr(spreadsheet, attribute, attr_value)
    # Make a lsit of all emails from the Active Members sheet
    # active_emails= query_active_member(spreadsheet, "email")
    # Send an email to each email in list
    # for email in active_emails:
    #     subject, body = generate_email(template, email, spreadsheet)
    #     send_email(email, subject, body)
    #     exit()

def main(file):
    """Prompts user for input. Prints out a main menu, asking
    if user would like to enter a new member into the spreadsheet, 
    change member information, 
    make a member inactive, or query the spreadsheet for information.
    Takes a command line argument of the file from which to read and edit."""
    while True:
        print("------------------------------------------")
        print("Welcome to the VOP Membership Portal!\n" \
            "What would you like to do?")
        print("[1] Add new members\n" \
            "[2] Update existing members\n" \
            "[3] Remove a member from Active Members\n" \
            "[4] Reinstate a member to Active Members\n" \
            "[5] Search for an existing member\n" \
            "[6] Create a nametag for an existing member\n" \
            "[7] Create a label sheet for an existing member\n" \
            "[8] Send an email to a mailing list\n" \
            "[9] Exit")
        print("------------------------------------------")
        response = input("Please enter a number from the main menu:\n")
        if response == "9":
            break
        elif response == "1":
            # Double check that the user made the right choice
            check = input("You've chosen to Add a new member. Proceed?\n" \
                          "(enter 'y' to continue, or 'n' to go back to main menu):\n")
            if check.lower() == "y":
                while True:
                    add_member(file)
                    # Check if the user wishes to add another member, or go back
                    # to main menu
                    again = input("Would you like to add another member? [y/n]\n")
                    if again.lower() == "y":
                        continue
                    else:
                        break
            else:
                continue
        elif response == "2":
            continue
        elif response == "3":
            continue
        elif response == "4":
            continue
        elif response == "5":
            continue
        elif response == "6":
            continue
        elif response == "7":
            continue
        elif response == "8":
            check = input("You have selected Send an email to a mailing list. Proceed? [y/n]\n")
            if check == "y":
                template = input("Please enter the file for the email template: \n")
                
                # all_active_email(template, file)
            else:
                continue       
        else:
            print("Not a valid response. Please enter a number from the main menu.")
            continue
            
if __name__ == '__main__':
    main("vopmembership_data 3.xlsx")