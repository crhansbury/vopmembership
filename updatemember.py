from openpyxl import load_workbook
from openpyxl.styles import Border, Side, Alignment
from datetime import datetime
from querymember import query_member_object
from emails import send_email, generate_email

def update_member():
    """Updates the member information in the spreadsheet. Prompts user for:
     name, pronouns, voice part, role, email address, phone number, or 
      mailing address."""
    pass

def inactive_member(file, attribute, attr_value):
    """Moves a member from the Active Member sheet to the Inactive Member 
    sheet."""
    member = query_member_object(file, attribute, attr_value)[0]
    # # Open the inactive sheet and append the member to the inactive sheet
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
