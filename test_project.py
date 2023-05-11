import unittest
from unittest.mock import patch
from datetime import datetime
from openpyxl import load_workbook
from project import add_member

class ProjectTest(unittest.TestCase):
    file = "vopmembership_data_unittest.xlsx"
    
    # Creating patches for input and internal functions
    @patch('project.send_email')
    @patch('project.create_member_id', return_value='999')
    @patch('builtins.input', side_effect=['John', 'Doe', 'he/him', 'T1', 'Member', 'vopmembershiptest+jdoe@gmail.com', '1234567890', '123 Street', 'City', 'CA', '12345'])
    def test_add_member(self, mock_input, mock_id, mock_send):
        # Call a test file to use as the spreadsheet
        file = "vopmembership_data_unittest.xlsx"
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

        # Check if email functions work properly
        mock_generate.assert_called_with(expected_template, expected_email, file)
        mock_send.assert_called_once()
        mock_send.assert_called_with(expected_email, expected_subject, expected_body)

        # Clean up the test file
        sheet.delete_rows(idx=40)
        workbook.save(file)

if __name__ == '__main__':
    unittest.main()
