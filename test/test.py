import os
import sys
sys.path.append(os.path.dirname(os.path.dirname(os.path.realpath(__file__))))
from Formatter import *

import unittest



class TestFormatter(unittest.TestCase):

    def _format_test_helper(self, f):
        """
        Helper function to test a format
        """
        format_dir = 'formats'
        if not os.path.exists("test") or not os.path.isdir("test"):
            if os.path.exists(os.path.join("..", format_dir)) and os.path.isdir(os.path.join("..", format_dir)):
                format_dir = os.path.join('..', format_dir)
        format_path = os.path.join(format_dir, f)

        data_dir = 'data'
        if os.path.exists("test") and os.path.isdir("test"):
            if os.path.exists(os.path.join("test", data_dir)) and os.path.isdir(os.path.join("test", data_dir)):
                data_dir = os.path.join('test', data_dir)

        input_fname = f + '.xlsx'
        input_path = os.path.join(data_dir, input_fname)
        input_fname_prefix, input_type = os.path.splitext(input_fname)

        actual_path = os.path.join(data_dir, input_fname_prefix + '_formatted.xlsx')
        expected_path = os.path.join(data_dir, input_fname_prefix + '_expected.xlsx')

        Formatter(FormatterOptions(format_path, input_path)).run()

        actual_SR = SpreadsheetReader(actual_path)
        expected_SR = SpreadsheetReader(expected_path)

        actual_file_data = actual_SR.get_rows()
        expected_file_data = expected_SR.get_rows()

        try:
            self.assertEqual(actual_file_data, expected_file_data)
        except Exception as e:
            print('Expected:\n', expected_file_data)
            print('Actual:\n', actual_file_data)
            print(e)



    def test_formats(self):
        """
        Test formats found in <root_dir>/formats
        
        Works by running the Formatter on the <format>.xlsx file which should be maintained in
        <root_dir>/test/data directory

        Also to be maintained is a 'expected' .xlsx file which should have the file format that the 
        Formatter is expected to create in the output xlsx file

        The output file is named <format>_formatted.xlsx

        Test passes if the user-maintained <format>_expected.xlsx is equivalent to
        program-generated <format>_formatted.xlsx

        Currently this is the way to test the custom operations in any formats
        """

        formats = [
                    'riceland'
                ]

        for f in formats:
            print('testing format ' + f)
            self._format_test_helper(f)



    def _rule_test_helper(self, rule):
        """
        Helper function to test a rule
        """

        rule_dir = 'rules'
        if os.path.exists("test") and os.path.isdir("test"):
            if os.path.exists(os.path.join("test", rule_dir)) and os.path.isdir(os.path.join("test", rule_dir)):
                rule_dir = os.path.join('test', rule_dir)
        rule_path = os.path.join(rule_dir, rule)
      
        data_dir = 'data'
        if os.path.exists("test") and os.path.isdir("test"):
            if os.path.exists(os.path.join("test", data_dir)) and os.path.isdir(os.path.join("test", data_dir)):
                data_dir = os.path.join('test', data_dir)

        input_fname = rule + '.xlsx'
        input_path = os.path.join(data_dir, input_fname)
        input_fname_prefix, input_type = os.path.splitext(input_fname)

        actual_path = os.path.join(data_dir, input_fname_prefix + '_formatted.xlsx')
        expected_path = os.path.join(data_dir, input_fname_prefix + '_expected.xlsx')

        Formatter(FormatterOptions(rule_path, input_path)).run()

        actual_SR = SpreadsheetReader(actual_path)
        expected_SR = SpreadsheetReader(expected_path)

        actual_file_data = actual_SR.get_rows()
        expected_file_data = expected_SR.get_rows()

        try:
            self.assertEqual(actual_file_data, expected_file_data)
        except Exception as e:
            print('Expected:\n', expected_file_data)
            print('Actual:\n', actual_file_data)
            print(e)



    def test_rules(self):
        """
        Test formats found in <root_dir>/test/rules
        
        Works by running the Formatter on the <rule>.xlsx file which should be maintained in
        <root_dir>/test/data directory

        Also to be maintained is a 'expected' .xlsx file which should have the file format that the 
        Formatter is expected to create in the output xlsx file

        The output file is named <rule>_formatted.xlsx

        Test passes if the user-maintained <rule>_expected.xlsx is equivalent to
        program-generated <rule>_formatted.xlsx
        """

        rules = [
                    'row_drop', 
                    'header_replace', 
                    'col_unique', 
                    'col_drop', 
                    'col_clear',
                    'col_replace_with_val',
                    'col_replace_based_on_col',
                    'col_cutpaste',
                    'col_new_col_replace',
                    'col_new_concat'
                ]

        for rule in rules:
            print('testing rule ' + rule)
            self._rule_test_helper(rule)


if __name__ == '__main__':
    unittest.main()
