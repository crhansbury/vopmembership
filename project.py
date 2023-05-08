from openpyxl import load_workbook
from openpyxl.styles import Border, Side, Alignment
from datetime import datetime
from emails import generate_email, send_email
from newmember import create_member_id
from querymember import query_active_member, query_member_object

def add_member(file):
    """Adds a member to the spreadsheet. Takes criteria for:
    name, pronouns, voice part, role, email address, phone number, and 
    mailing address. Appends spreadsheet with this information. Updates
    a field that indicates the date of the latest update."""
    # Prompt the user for info
    print("Please enter the information for the new member.")
    first_name = input("First Name: ").title()
    last_name = input("Last Name: ").title()
    pronouns = input("Pronouns: ").lower()
    section = input("Section: ").upper()
    role = input("Role: ").title()
    email = input("Email: ")
    phone = input("Phone Number: ")
    address = input("Street Address: ").title()
    city = input("City: ").title()
    state = input("State: ").upper()
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
    row = (
        last_name,first_name,pronouns,section,member_id,role,address,city, \
        state,zipcode,phone,email
    )
    sheet.append(row)
    # Formatting the new row to match the rest of the spreadsheet
    thin_border = Border(left=Side(style='thin'),
                         right=Side(style='thin'),
                         top=Side(style='thin'),
                         bottom=Side(style='thin'))
    alignment = Alignment(horizontal='left', vertical='bottom', wrap_text=True)
    new_row = sheet.max_row
    for cell in sheet[new_row]:
        cell.border = thin_border
        cell.alignment = alignment
    # Adding the date modified
    date = datetime.now()
    sheet["M1"] = "Date Modified: {} at {}".format(date.strftime("%m/%d/%Y"),
                                                   date.strftime("%H:%M"))
    workbook.save(file)
    print("{} {} successfully added to {}.".format(first_name, last_name, file))
    # Generating and sending welcome email once new member has been successfully
    # appended to spreadsheet
    try:
        subject, body = generate_email("email-templates/new_member_template.txt", 
                                       email, file)
        send_email(email, subject, body)
    except:
        print("Email to {} was unsuccessful.".format(email))

def all_active_email(template, spreadsheet):
    """Sends an email to every member in the Active Members sheet of the 
    spreadsheet passed into the 'spreadsheet' parameter. Uses the email
    template passed into the 'template' parameter."""
    # Make a lsit of all emails from the Active Members sheet
    active_emails= query_active_member(spreadsheet, "email")
    # Send an email to each email in list
    for email in active_emails:
        subject, body = generate_email(template, email, spreadsheet)
        send_email(email, subject, body)

def inactive_member(file, attribute, attr_value):
    """Moves a member from the Active Member sheet to the Inactive Member 
    sheet."""
    member = query_member_object(file, attribute, attr_value)[0]
    # Open the inactive sheet and append the member to the inactive sheet
    workbook = load_workbook(file)
    inactive_sheet = workbook["Inactive Members"]
    active_sheet = workbook["Active Members"]
    # Find the row that matches the email attribute
    active_sheet_row = 0
    for row_num, row in enumerate(active_sheet.iter_rows(min_row=2, min_col=12, 
                                                         max_col=12, 
                                                         values_only=True), 
                                                         start=2):
        if member.email == row[0]:
            active_sheet_row = row_num
    # Find the last row of inactive sheet
    inactive_sheet_row = inactive_sheet.max_row + 1
    # Copy/pasting all values from active sheet to inactive sheet
    for col_num, cell in enumerate(active_sheet[active_sheet_row], 1):
        value = cell.value
        inactive_sheet.cell(row=inactive_sheet_row, column=col_num, value=value)
    # Formatting the cells
    thin_border = Border(left=Side(style='thin'),
                         right=Side(style='thin'),
                         top=Side(style='thin'),
                         bottom=Side(style='thin'))
    alignment = Alignment(horizontal='left', vertical='bottom', wrap_text=True)
    for cell in inactive_sheet[inactive_sheet_row]:
        cell.border = thin_border
        cell.alignment = alignment
    # Delete the member from active member sheet
    active_sheet.delete_rows(idx=active_sheet_row)
    # Add date modified
    date = datetime.now()
    active_sheet["M1"] = "Date Modified: {} at {}".format(date.strftime("%m/%d/%Y"),
                                                   date.strftime("%H:%M"))
    workbook.save(file)
    print("{} successfully moved to Inactive Members.".format(member.first_name))
    # Send an automatic email when this completes
    subject, body = generate_email("email-templates/inactive_member_template.txt",
                                   member.email, file)
    send_email(member.email, subject, body)

def active_member(file, attribute, attr_value):
    """Moves a member from the Inactive Member sheet to the Active Member 
    sheet. Takes the argument of a member attribute and value, and looks
    up that member. Copies the member over from the Inactive sheet to the 
    Active sheet."""
    member = query_member_object(file, attribute, attr_value)[0]
    # Open the active sheet and append the member to the active sheet
    workbook = load_workbook(file)
    inactive_sheet = workbook["Inactive Members"]
    active_sheet = workbook["Active Members"]
    # Find the row that matches the email attribute
    inactive_sheet_row = 0
    for row_num, row in enumerate(inactive_sheet.iter_rows(min_row=2, min_col=12, 
                                                         max_col=12, 
                                                         values_only=True), 
                                                         start=2):
        if member.email == row[0]:
            inactive_sheet_row = row_num
    # Find the last row of inactive sheet
    active_sheet_row = active_sheet.max_row + 1
    # Copy/pasting all values from inactive sheet to active sheet
    for col_num, cell in enumerate(inactive_sheet[inactive_sheet_row], 1):
        value = cell.value
        active_sheet.cell(row=active_sheet_row, column=col_num, value=value)
    # Formatting the cells
    thin_border = Border(left=Side(style='thin'),
                         right=Side(style='thin'),
                         top=Side(style='thin'),
                         bottom=Side(style='thin'))
    alignment = Alignment(horizontal='left', vertical='bottom', wrap_text=True)
    for cell in active_sheet[active_sheet_row]:
        cell.border = thin_border
        cell.alignment = alignment
    # Delete member from inactive member sheet
    inactive_sheet.delete_rows(idx=inactive_sheet_row)
    # Add date modified
    date = datetime.now()
    active_sheet["M1"] = "Date Modified: {} at {}".format(date.strftime("%m/%d/%Y"),
                                                   date.strftime("%H:%M"))
    workbook.save(file)
    print("{} successfully moved to Active Members.".format(member.first_name))
    # Send an automatic email when this completes
    subject, body = generate_email("email-templates/active_member_template.txt",
                                   member.email, file)
    send_email(member.email, subject, body)

def search_member(file, attribute, attr_value):
    """Searches the spreadsheet for members that match the search criteria.
    Prints all the information for the members who match the search.
    Numbers all of the results if there are multiple members returned. Uses
    the query_member_object() function to search the file."""
    query_list = query_member_object(file, attribute, attr_value)
    if len(query_list) < 1:
        print("No members meet your search criteria.")
    else:
        if len(query_list) > 1:
            print("Multiple members meet your search criteria.")
        member_count = 1
        for member in query_list:
            print("Result number {}:".format(member_count))
            print("Name: {} {}\n".format(member.first_name, member.last_name),
                "Pronouns: {}\n".format(member.pronouns),
                "Section: {}\n".format(member.section),
                "Member ID: {}\n".format(member.id),
                "Email: {}\n".format(member.email),
                "Phone number: {}\n".format(member.phone),
                "Address: {address}, {city} {state} {zip}".format(address=member.address,
                                                                    city=member.city,
                                                                    state=member.state,
                                                                    zip=member.zip))
            member_count += 1

def update_member(file, email):
    """Updates the member information in the spreadsheet. Prompts user for:
     name, pronouns, voice part, role, email address, phone number, or 
      mailing address."""
    # Open the file and iterate over the emails field to find the desired row
    workbook = load_workbook(file)
    sheet = workbook["Active Members"]
    edit_row = 0
    best_email = email
    for row_num, row in enumerate(sheet.iter_rows(min_row=2, min_col=1, 
                                                  max_col=12, values_only=True), 
                                                  start=2):
        if email == row[11]:
            edit_row = row_num
            member_first = row[1]
            member_last = row[0]
    if edit_row == 0:
        print("{} not found in {}.".format(email, file))
    else:
        check = input("Edit information for {} {}? [y/n]\n".format(member_first, member_last))
        if check.lower().strip() == "y":
            while True:
                print("Please select which field you would like to edit\n",
                      "[1] First Name\n",
                      "[2] Last Name\n",
                      "[3] Pronouns\n",
                      "[4] Section\n",
                      "[5] Email\n",
                      "[6] Phone number\n",
                      "[7] Address\n",
                      "[8] Role")
                response = input("(enter a number from the menu above)\n")
                if response == "1":
                    first_name = input("Please enter new first name: ").title()
                    sheet.cell(row=edit_row, column=2, value=first_name)
                    member_first = first_name
                    # Prompt user for another edit
                    again = input("Edit another field? [y/n]\n").lower().strip()
                    if again == "y":
                        continue
                    else:
                        break
                if response == "2":
                    last_name = input("Please enter new last name: ").title()
                    sheet.cell(row=edit_row, column=1, value=last_name)
                    member_last = last_name
                    # Prompt user for another edit
                    again = input("Edit another field? [y/n]\n").lower().strip()
                    if again == "y":
                        continue
                    else:
                        break
                if response == "3":
                    pronouns = input("Please enter new pronouns: ").lower()
                    sheet.cell(row=edit_row, column=3, value=pronouns)
                    # Prompt user for another edit
                    again = input("Edit another field? [y/n]\n").lower().strip()
                    if again == "y":
                        continue
                    else:
                        break
                if response == "4":
                    section = input("Please enter new section: ").upper()
                    sheet.cell(row=edit_row, column=4, value=section)
                    # Prompt user for another edit
                    again = input("Edit another field? [y/n]\n").lower().strip()
                    if again == "y":
                        continue
                    else:
                        break
                if response == "5":
                    new_email = input("Please enter new email: ")
                    sheet.cell(row=edit_row, column=12, value=new_email)
                    best_email = new_email
                    # Prompt user for another edit
                    again = input("Edit another field? [y/n]\n").lower().strip()
                    if again == "y":
                        continue
                    else:
                        break
                if response == "6":
                    phone = input("Please enter new phone number: ")
                    sheet.cell(row=edit_row, column=11, value=phone)
                    # Prompt user for another edit
                    again = input("Edit another field? [y/n]\n").lower().strip()
                    if again == "y":
                        continue
                    else:
                        break
                if response == "7":
                    street_address = input("Please enter new street address: ").title()
                    sheet.cell(row=edit_row, column=7, value=street_address)
                    city = input("Please enter new city: ").title()
                    sheet.cell(row=edit_row, column=8, value=city)
                    state = input("Please enter new state: ").upper()
                    sheet.cell(row=edit_row, column=9, value=state)
                    zipcode = input("Please enter new zip code: ")
                    sheet.cell(row=edit_row, column=10, value=zipcode)
                    # Prompt user for another edit
                    again = input("Edit another field? [y/n]\n").lower().strip()
                    if again == "y":
                        continue
                    else:
                        break
                if response == "8":
                    role = input("Please enter new role: ").title()
                    sheet.cell(row=edit_row, column=6, value=role)
                    # Prompt user for another edit
                    again = input("Edit another field? [y/n]\n").lower().strip()
                    if again == "y":
                        continue
                    else:
                        break
                else:
                    print("Please enter a valid response from the menu.")
                    continue
            # Adding the date modified
            date = datetime.now()
            sheet["M1"] = "Date Modified: {} at {}".format(date.strftime("%m/%d/%Y"),
                                                   date.strftime("%H:%M"))
            workbook.save(file)
            print("{} {} successfully edited.".format(member_first, member_last))
            subject, body = generate_email("email-templates/updated_member_template.txt",
                                           best_email, file)
            send_email(best_email, subject, body)

def main(file):
    """Prompts user for input. Prints out a main menu, asking
    if user would like to enter a new member into the spreadsheet, 
    change member information, 
    make a member inactive, or query the spreadsheet for information.
    Takes a command line argument of the file from which to read and edit."""
    # Test the file at the beginning to ensure it's the right format to avoid
    # testing it for every function
    while True:
        print("------------------------------------------")
        print("Welcome to the VOP Membership Portal!\n" \
              "What would you like to do?")
        print("[1] Add new members\n" \
            "[2] Remove members from Active Members\n" \
            "[3] Reinstate members to Active Members\n" \
            "[4] Update active members\n" \
            "[5] Search for existing members\n" \
            "[6] Create a nametag for an existing member\n" \
            "[7] Create a label sheet for an existing member\n" \
            "[8] Send an email to all active members\n" \
            "[9] Exit")
        print("------------------------------------------")
        response = input("Please enter a number from the main menu:\n")
        if response == "9":
            break
        elif response == "1":
            # Double check that the user made the right choice
            check = input("You have selected 'Add a new member'. Proceed?\n" \
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
            check = input("You have selected 'Remove a member from Active " \
                          "Members'. Proceed? [y/n]\n")
            if check.lower().strip() == "y":
                while True:
                    while True:
                        print("Please choose one of the following types of "\
                            "information for the member you wish you remove:\n"\
                            "[1] Email\n"\
                            "[2] Member ID")
                        response = input("")
                        attribute = ""
                        if response == "1":
                            attribute = "email"
                            break
                        elif response == "2":
                            attribute = "id"
                            break
                        else:
                            print("Please select a valid item from the menu.")
                            continue
                    attr_value = input("Please enter the member's {}:\n".format(attribute))
                    try:
                        inactive_member(file, attribute, attr_value)
                    except: 
                        print("{} is not a valid {}.".format(attr_value, attribute))
                        continue
                    again = input("Would you like to remove another member? [y/n]\n")
                    if again.lower().strip() == "y":
                        continue
                    else:
                        break
            else:
                continue
        elif response == "3":
            check = input("You have selected 'Reinstate a member to Active "\
                          "Members'. Proceed? [y/n]\n")
            if check.lower().strip() == "y":
                while True:
                    while True:
                        print("Please choose one of the following types of " 
                              "information for the member you wish to " \
                              "reinstate:\n" \
                              "[1] Email\n"\
                              "[2] Member ID")
                        response = input("")
                        attribute = ""
                        if response == "1":
                            attribute = "email"
                            break
                        elif response == "2":
                            attribute = "id"
                            break
                        else:
                            print("Please select a valid item from the menu.")
                            continue
                    attr_value = input("Please enter the member's {}:\n".format(attribute))
                    try:
                        active_member(file, attribute, attr_value)
                    except: 
                        print("{} is not a valid {}.".format(attr_value, attribute))
                        continue
                    again = input("Would you like to reinstate another member? [y/n]\n")
                    if again.lower().strip() == "y":
                        continue
                    else:
                        break
            else:
                continue
        elif response == "4":
            check = input("You have selected 'Update active members'. "\
                          "Proceed? [y/n]\n")
            if check.lower().strip() == "y":
                while True:
                    email = input("Please enter the email for the member whose "\
                                  "information you wish you update\n")
                    update_member(file, email)
                    again = input("Update another member? [y/n]\n").lower().strip()
                    if again == "y":
                        continue
                    else:
                        break
            else:
                continue
        elif response == "5":
            check = input("You have selected 'Search for an existing member'. "\
                          "Proceed? [y/n]\n")
            if check.lower().strip() == "y":
                while True:
                    print("Please choose the search criteria:\n",
                          "[1] First Name\n",
                          "[2] Last Name\n",
                          "[3] Pronouns\n",
                          "[4] Section\n",
                          "[5] Role\n",
                          "[6] Member ID\n",
                          "[7] Email\n",
                          "[8] Phone number\n",
                          "[9] Street Address\n",
                          "[10] City\n",
                          "[11] State\n",
                          "[12] Zip Code")
                    reply = input("(choose a number from the menu above)\n")
                    attribute = ""
                    attr_value = ""
                    if reply == "1":
                        attribute = "first_name"
                        attr_value = input("Please enter the member's first name:\n").title()
                    elif reply == "2":
                        attribute = "last_name"
                        attr_value = input("Please enter the member's last name:\n").title()                     
                    elif reply == "3":
                        attribute = "pronouns"
                        attr_value = input("Please enter the member's pronouns:\n").lower()                      
                    elif reply == "4":
                        attribute = "section"
                        attr_value = input("Please enter the member's section:\n").upper()
                    elif reply == "5":
                        attribute = "role"
                        attr_value = input("Please enter the member's role:\n").title()
                    elif reply == "6":
                        attribute = "id"
                        attr_value = input("Please enter the member's member ID:\n")                       
                    elif reply == "7":
                        attribute = "email"
                        attr_value = input("Please enter the member's email:\n")                        
                    elif reply == "8":
                        attribute = "phone"
                        attr_value = input("Please enter the member's phone number:\n")                        
                    elif reply == "9":
                        attribute = "address"
                        attr_value = input("Please enter the member's street address:\n").title()                        
                    elif reply == "10":
                        attribute = "city"
                        attr_value = input("Please enter the member's city:\n").title()                        
                    elif reply == "11":
                        attribute = "state"
                        attr_value = input("Please enter the member's state:\n").upper()                        
                    elif reply == "12":
                        attribute = "zip"
                        try:
                            attr_value = int(input("Please enter the member's zip code:\n"))
                        except ValueError:
                            print("Not a valid zip code.")
                            continue                        
                    else:
                        print("Please enter a valid number from the menu.")
                        continue
                    search_member(file, attribute, attr_value)
                    again = input("Would you like to perform another search? [y/n]\n")
                    if again.lower().strip() == "y":
                        continue
                    else:
                        break
            else:
                continue
        elif response == "6":
            continue
        elif response == "7":
            continue
        elif response == "8":
            check = input("You have selected 'Send an email to all active " \
                          "members'. Proceed? [y/n]\n")
            if check == "y":
                template = input("Please enter the email template file: \n")
                try:
                    all_active_email(template, file)
                    continue
                except:
                    print("invalid file")
                    continue
            else:
                continue       
        else:
            print("Not a valid response. Please enter a number from the main menu.")
            continue
            
if __name__ == '__main__':
    main("vopmembership_data 5.xlsx")