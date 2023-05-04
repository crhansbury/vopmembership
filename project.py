from openpyxl import load_workbook
from datetime import datetime
from emails import new_member_email, send_email
from newmember import create_member_id

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
        subject, body = new_member_email("email-templates/new_member_template.txt", email, file)
        send_email(email, subject, body)
    except:
        print("Email to {} was unsuccessful.".format(email))

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
            "[8] Exit")
        print("------------------------------------------")
        response = input("Please enter a number from the main menu): ")
        if response == "8":
            break
        elif response == "1":
            # Double check that the user made the right choice
            check = input("You've chosen to Add a new member. Proceed?\n" \
                          "(enter Y to continue, or N to go back to main menu:\n")
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
        else:
            print("Not a valid response. Please enter a number from the main menu.")
            continue
            
if __name__ == '__main__':
    main("vopmembership_data 3.xlsx")