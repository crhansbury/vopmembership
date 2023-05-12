import unittest
from unittest.mock import patch, call
from io import StringIO
from datetime import datetime
from classes import Member
from openpyxl import load_workbook
from project import add_member, all_active_email, search_member, main

class ProjectTest(unittest.TestCase):

    # Patch decorators for input, stdout, and internal functions: create_member_id() and
    # send_email()
    @patch('sys.stdout', new_callable=StringIO)
    @patch('project.send_email')
    @patch('project.create_member_id', return_value='999')
    @patch('builtins.input', side_effect=['John', 'Doe', 'he/him', 'T1', 'Member', 'vopmembershiptest+jdoe@gmail.com', '1234567890', '123 Street', 'City', 'CA', '12345'])
    def test_add_member(self, mock_input, mock_id, mock_send, mock_stdout):
        # Call a test file to use as the spreadsheet
        file = "data-files/vopmembership_data_unittest.xlsx"
        expected_template = 'email-templates/new_member_template.txt'
        expected_email = 'vopmembershiptest+jdoe@gmail.com'
        expected_subject = 'Welcome to VOP, John!'
        expected_body = """Dear John,

Welcome to VOP! We are so excited for you to join us this semester. You will be singing with us in the T1 section, and your section leader will be Jerome Munsch. Please reach out to Jerome at vopmembershiptest+jmunsch@gmail.com if you have any questions!

Just to ensure we have all your information correct, please review the following:

Name: John Doe
Pronouns: he/him
Section: T1
Member ID: 999
Email: vopmembershiptest+jdoe@gmail.com
Phone Number: 1234567890
Address: 123 Street, City CA 12345

If anything is incorrect please reply to this email. 
We look forward to seeing you at our first rehearsal!

Best,
VOP Membership team"""
        # Setting up a mock generate_email() function
        with patch('project.generate_email') as mock_generate:
            mock_generate.return_value = (expected_subject, expected_body)
            add_member(file)
        
        workbook = load_workbook(filename=file)
        sheet = workbook["Active Members"]
        # Check if the member was added correctly
        self.assertEqual(sheet['A40'].value, 'Doe')
        self.assertEqual(sheet['B40'].value, 'John')
        self.assertEqual(sheet['C40'].value, 'he/him')
        self.assertEqual(sheet['D40'].value, 'T1')
        self.assertEqual(sheet['E40'].value, '999')
        self.assertEqual(sheet['F40'].value, 'Member')
        self.assertEqual(sheet['G40'].value, '123 Street')
        self.assertEqual(sheet['H40'].value, 'City')
        self.assertEqual(sheet['I40'].value, 'CA')
        self.assertEqual(sheet['J40'].value, '12345')
        self.assertEqual(sheet['K40'].value, '1234567890')
        self.assertEqual(sheet['L40'].value, 'vopmembershiptest+jdoe@gmail.com')
        self.assertEqual(sheet['M40'].value, 'Active')

        # Check if the date modified was updated
        date = datetime.now()
        expected_date = '{} at {}'.format(date.strftime("%m/%d/%Y"), date.strftime("%H:%M"))
        self.assertEqual(sheet['N2'].value, expected_date)

        # Check stdout to make sure the print statements are correct
        expected_stdout = "Please enter the information for the new member.\n✅ John Doe successfully added to data-files/vopmembership_data_unittest.xlsx.\n"
        self.assertEqual(mock_stdout.getvalue(), expected_stdout)
        
        # Check if email functions work properly
        mock_generate.assert_called_with(expected_template, expected_email, file)
        mock_send.assert_called_once()
        mock_send.assert_called_with(expected_email, expected_subject, expected_body)

        # Clean up the test file
        sheet.delete_rows(idx=40)
        workbook.save(file)
        workbook.close()

    # Patch decorators for mock query_active_member, generate_email and 
    # send_email functions
    @patch('project.send_email')
    @patch('project.generate_email')
    @patch('project.query_active_member')
    def test_all_active_email(self, mock_query, mock_generate, mock_send):
        # Mocking the dependencies for testing
        template = "email-templates/test_template.txt"
        file = "data-files/vopmembership_data_unittest.xlsx"

        # Define the list of emails to iterate over as the return value of query_active_members
        mock_query.return_value = ['vopmembershiptest+test1@gmail.com', 
                                   'vopmembershiptest+test2@gmail.com', 
                                   'vopmembershiptest+test3@gmail.com']

        # Set the side effect of mock_generate_email to return different values for each call
        mock_generate.side_effect = lambda tmpl, email, fl: ('Test Subject', f'Test Body for {email}')

        # Call the function
        all_active_email(template, file)

        # Assert that generate_email() was called with the expected emails and check the return values
        calls = [call(template, email, file) for email in mock_query.return_value]
        mock_generate.assert_has_calls(calls)
        mock_generate.assert_called_with(template, mock_query.return_value[-1], file)

        # Assert the number of generate_email() function calls
        self.assertEqual(mock_generate.call_count, len(mock_query.return_value))

        # Assert that send_email was called as many times as there are emails
        self.assertEqual(mock_send.call_count, len(mock_query.return_value))

    # Patch decorators for a mock print and query_member_object()
    @patch('builtins.print')
    @patch('project.query_member_object')
    def test_search_member(self, mock_query, mock_print):
        file = "data-files/vopmembership_data_unittest.xlsx"
        attribute = 'id'
        attr_value = '1'
        
        # Set the return value for mock_query
        mock_query.return_value = [Member(id='1', 
                                          first_name='Johnson', 
                                          last_name='White', 
                                          pronouns='they/them', 
                                          section='B1', 
                                          role=None, 
                                          email='vopmembershiptest+jwhite@gmail.com', 
                                          address='10932 Bigge Rd', 
                                          city='Menlo Park', 
                                          state='CA', 
                                          zip=94025, 
                                          phone='408 496-7223', 
                                          status='Active')]
        
        # Call the function
        search_member(file, attribute, attr_value)

        # Assert that the search is being called with expected values
        mock_query.assert_has_calls([call(file, attribute, attr_value)])

        # Assert that the printed output is correct
        expected_calls = [
            call("1 member(s) meet your search criteria."),
            call("⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽"),
            call("Result number 1:"),
            call(' Name: Johnson White\n', 
                 'Pronouns: they/them\n', 
                 'Section: B1\n', 
                 'Role: None\n', 
                 'Member ID: 1\n', 
                 'Email: vopmembershiptest+jwhite@gmail.com\n', 
                 'Phone number: 408 496-7223\n', 
                 'Address: 10932 Bigge Rd, Menlo Park CA 94025\n', 
                 'Status: Active')
        ]
        mock_print.assert_has_calls(expected_calls)
    
    # Patch decorators for mock_input and mock_stdout
    @patch('sys.stdout', new_callable=StringIO)
    @patch('builtins.input')
    def test_main(self, mock_input, mock_stdout):
        file = 'data-files/vopmembership_data_unittest.xlsx'

        # Define mock_input to check that main menu works as expected
        mock_input.side_effect= ['1', 'n', '7']
        
        # Call the main function
        main(file)

        # Check stdout for expected output
        expected_stdout = "⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥\n"\
                          "   Welcome to the VOP Membership Portal!\n" \
                          "   What would you like to do?\n" \
                          "🟥 [1] Add new members\n" \
                          "🟧 [2] Remove members from Active Members\n" \
                          "🟨 [3] Reinstate members to Active Members\n" \
                          "🟩 [4] Update active members\n" \
                          "🟦 [5] Search for member information\n" \
                          "🟪 [6] Send an email to all active members\n" \
                          "⬛️ [7] Exit\n" \
                          "⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥\n" \
                          "⬅️  Returning to Main Menu.\n" \
                          "⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥\n"\
                          "   Welcome to the VOP Membership Portal!\n" \
                          "   What would you like to do?\n" \
                          "🟥 [1] Add new members\n" \
                          "🟧 [2] Remove members from Active Members\n" \
                          "🟨 [3] Reinstate members to Active Members\n" \
                          "🟩 [4] Update active members\n" \
                          "🟦 [5] Search for member information\n" \
                          "🟪 [6] Send an email to all active members\n" \
                          "⬛️ [7] Exit\n" \
                          "⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥⪥\n" \
                          "👋 Goodbye!\n"
        self.assertEqual(mock_stdout.getvalue(), expected_stdout)

        # Check for errors with bad files
        bad_sheet_file = 'data-files/unittest_badsheet.xlsx' # No Inactive Members Sheet
        with self.assertRaises(SystemExit) as bad_sheet:
            main(bad_sheet_file)
        # Assert the correct error code
        self.assertEqual(bad_sheet.exception.code, 2)
        bad_column_file = 'data-files/unittest_badcolumns.xlsx' # Wrong column names
        with self.assertRaises(SystemExit) as bad_columns:
            main(bad_column_file)
        self.assertEqual(bad_columns.exception.code, 3)
        bad_exists_file = 'nofile.xlsx' # Nonexistant file
        with self.assertRaises(SystemExit) as bad_exists:
            main(bad_exists_file)
        self.assertEqual(bad_exists.exception.code, 1)


if __name__ == '__main__':
    unittest.main()
