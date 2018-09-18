#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Tests for Transtab."""

import os

import unittest
import tablib

from transtab import TransTab


class TablibTestCase(unittest.TestCase):
    """Tablib test cases."""

    def setUp(self):
        from warnings import filterwarnings
        filterwarnings("ignore", category=UserWarning)
        filterwarnings("ignore", category=DeprecationWarning)

    def tearDown(self):
        pass


    def create_data_file(self, data, fname):
        fname_prefix, f_type = os.path.splitext(fname)
        f_type = f_type.lstrip('.')
        with open(fname, 'wb') as f:
            f.write(data.export(f_type))


    def create_format_file(self, test_command, fname):
        with open(fname, 'w') as f:
            f.write(test_command)


    def perform_common_steps(self, test_command, in_data, exp_data):
        format_code = test_command.split()[0]
        in_fname = 'in_'+ format_code +'.xlsx'

        self.create_data_file(in_data, in_fname)

        format_fname = format_code + '.txt'
        self.create_format_file(test_command, format_fname)

        act_fname = 'act_' + format_code + '.xlsx'

        TransTab(in_fname=in_fname, format_fname=format_fname, out_fname=act_fname).transform()

        with open(act_fname, 'rb') as f:
            act_data = tablib.Dataset().load(f.read())

        self.assertEqual(act_data.dict, exp_data.dict)

        os.remove(in_fname)
        os.remove(act_fname)
        os.remove(format_fname)


    def test_drop(self):
        test_command = 'drop'

        headers = ('item', 'type', 'price')
        orange = ('Orange', 'Fruit', '90')
        cucumber = ('Cucumber', 'Vegetable', '67')
        pen = ('Pen', 'Stationery', '50')

        in_data = tablib.Dataset(headers=headers)
        in_data.append(orange)
        in_data.append(cucumber)
        in_data.append(pen)

        exp_data = tablib.Dataset(headers=headers)
        exp_data.append(orange)
        exp_data.append(cucumber)

        self.perform_common_steps(test_command, in_data, exp_data)

    def test_delete_rows_by_column_val(self):
        test_command = "delete-rows-by-column-val col 'item' val 'Cucumber'"

        headers = ('item', 'type', 'price')
        orange = ('Orange', 'Fruit', '90')
        cucumber = ('Cucumber', 'Vegetable', '67')
        pen = ('', 'Stationery', '50')

        in_data = tablib.Dataset(headers=headers)
        in_data.append(orange)
        in_data.append(cucumber)
        in_data.append(pen)

        exp_data = tablib.Dataset(headers=headers)
        exp_data.append(cucumber)
        exp_data.append(pen)

        self.perform_common_steps(test_command, in_data, exp_data)


    def test_new_col(self):
        test_command = "new col 'quantity'"

        headers = ('item', 'type', 'price')
        orange = ('Orange', 'Fruit', '90')
        cucumber = ('Cucumber', 'Vegetable', '67')

        in_data = tablib.Dataset(headers=headers)
        in_data.append(orange)
        in_data.append(cucumber)

        headers = ('item', 'type', 'price', 'quantity')
        orange = ('Orange', 'Fruit', '90', '')
        cucumber = ('Cucumber', 'Vegetable', '67', '')

        exp_data = tablib.Dataset(headers=headers)
        exp_data.append(orange)
        exp_data.append(cucumber)

        self.perform_common_steps(test_command, in_data, exp_data)


    def test_clear_col(self):
        test_command = "clear 'price'"

        headers = ('item', 'type', 'price')
        orange = ('Orange', 'Fruit', '90')
        cucumber = ('Cucumber', 'Vegetable', '67')

        in_data = tablib.Dataset(headers=headers)
        in_data.append(orange)
        in_data.append(cucumber)

        headers = ('item', 'type', 'price')
        orange = ('Orange', 'Fruit', '')
        cucumber = ('Cucumber', 'Vegetable', '')

        exp_data = tablib.Dataset(headers=headers)
        exp_data.append(orange)
        exp_data.append(cucumber)

        self.perform_common_steps(test_command, in_data, exp_data)


    def test_delete_col(self):
        test_command = "delete 'price'"

        headers = ('item', 'type', 'price')
        orange = ('Orange', 'Fruit', '90')
        cucumber = ('Cucumber', 'Vegetable', '67')

        in_data = tablib.Dataset(headers=headers)
        in_data.append(orange)
        in_data.append(cucumber)

        headers = ('item', 'type')
        orange = ('Orange', 'Fruit')
        cucumber = ('Cucumber', 'Vegetable')

        exp_data = tablib.Dataset(headers=headers)
        exp_data.append(orange)
        exp_data.append(cucumber)

        self.perform_common_steps(test_command, in_data, exp_data)


    def test_delete_row(self):

        test_command = 'delete row 3'

        headers = ('item', 'type', 'price')
        orange = ('Orange', 'Fruit', '90')
        cucumber = ('Cucumber', 'Vegetable', '67')
        pen = ('Pen', 'Stationery', '50')

        in_data = tablib.Dataset(headers=headers)
        in_data.append(orange)
        in_data.append(cucumber)
        in_data.append(pen)

        exp_data = tablib.Dataset(headers=headers)
        exp_data.append(orange)
        exp_data.append(pen)

        self.perform_common_steps(test_command, in_data, exp_data)


    def test_rename_col(self):

        test_command = "rename 'price' 'cost'"

        headers = ('item', 'type', 'price')
        orange = ('Orange', 'Fruit', '90')
        cucumber = ('Cucumber', 'Vegetable', '67')

        in_data = tablib.Dataset(headers=headers)
        in_data.append(orange)
        in_data.append(cucumber)

        headers = ('item', 'type', 'cost')

        exp_data = tablib.Dataset(headers=headers)
        exp_data.append(orange)
        exp_data.append(cucumber)

        self.perform_common_steps(test_command, in_data, exp_data)


    def test_copy_col(self):

        test_command = "copy 'price' 'cost'"

        headers = ('item', 'type', 'price', 'cost')
        orange = ('Orange', 'Fruit', '90', '')
        cucumber = ('Cucumber', 'Vegetable', '67', '')

        in_data = tablib.Dataset(headers=headers)
        in_data.append(orange)
        in_data.append(cucumber)

        orange = ('Orange', 'Fruit', '90', '90')
        cucumber = ('Cucumber', 'Vegetable', '67', '67')

        exp_data = tablib.Dataset(headers=headers)
        exp_data.append(orange)
        exp_data.append(cucumber)

        self.perform_common_steps(test_command, in_data, exp_data)


    def test_cutpaste_col(self):

        test_command = "cutpaste 'price' 'cost'"

        headers = ('item', 'type', 'price', 'cost')
        orange = ('Orange', 'Fruit', '90', '')
        cucumber = ('Cucumber', 'Vegetable', '67', '')

        in_data = tablib.Dataset(headers=headers)
        in_data.append(orange)
        in_data.append(cucumber)

        orange = ('Orange', 'Fruit', '', '90')
        cucumber = ('Cucumber', 'Vegetable', '', '67')

        exp_data = tablib.Dataset(headers=headers)
        exp_data.append(orange)
        exp_data.append(cucumber)

        self.perform_common_steps(test_command, in_data, exp_data)


    def test_concatenate_col(self):
        test_command = "concatenate 'item', 'type' and store in 'itemtype'"

        headers = ('item', 'type', 'price')
        orange = ('Orange', 'Fruit', '90')
        cucumber = ('Cucumber', 'Vegetable', '67')

        in_data = tablib.Dataset(headers=headers)
        in_data.append(orange)
        in_data.append(cucumber)

        headers = ('item', 'type', 'price', 'itemtype')
        orange = ('Orange', 'Fruit', '90', 'OrangeFruit')
        cucumber = ('Cucumber', 'Vegetable', '67', 'CucumberVegetable')

        exp_data = tablib.Dataset(headers=headers)
        exp_data.append(orange)
        exp_data.append(cucumber)

        self.perform_common_steps(test_command, in_data, exp_data)


    def test_replace_col(self):

        test_command = "replace 'item' {'Orange': 'OR'} case-insensitive default 'CODE'"

        headers = ('item', 'type', 'price')
        orange = ('Orange', 'Fruit', '90')
        cucumber = ('Cucumber', 'Vegetable', '67')

        in_data = tablib.Dataset(headers=headers)
        in_data.append(orange)
        in_data.append(cucumber)

        orange = ('OR', 'Fruit', '90')
        cucumber = ('CODE', 'Vegetable', '67')

        exp_data = tablib.Dataset(headers=headers)
        exp_data.append(orange)
        exp_data.append(cucumber)

        self.perform_common_steps(test_command, in_data, exp_data)


    def test_delete_duplicate_rows(self):

        test_command = "delete-duplicate-rows"

        headers = ('item', 'type', 'price')
        orange = ('Orange', 'Fruit', '90')
        cucumber = ('Cucumber', 'Vegetable', '67')

        in_data = tablib.Dataset(headers=headers)
        in_data.append(orange)
        in_data.append(orange)
        in_data.append(cucumber)

        exp_data = tablib.Dataset(headers=headers)
        exp_data.append(orange)
        exp_data.append(cucumber)

        self.perform_common_steps(test_command, in_data, exp_data)


    def test_sum_col_and_delete_duplicate_rows(self):

        test_command = "sum-col-and-delete-duplicate-rows sum 'price' unique 'item'"

        headers = ('item', 'type', 'price')
        orange = ('Orange', 'Fruit', 90)
        cucumber = ('Cucumber', 'Vegetable', 67)

        in_data = tablib.Dataset(headers=headers)
        in_data.append(orange)
        in_data.append(orange)
        in_data.append(cucumber)

        orange = ('Orange', 'Fruit', 180)

        exp_data = tablib.Dataset(headers=headers)
        exp_data.append(orange)
        exp_data.append(cucumber)

        self.perform_common_steps(test_command, in_data, exp_data)


    def test_validate_phone_number(self):

        test_command = "do validate_phone_number on 'Ph num'"

        headers = ('Name', 'Ph num')
        john = ('John', '')
        george = ('George', 'a')
        binoy = ('Binoy', '123')
        jack = ('Jack', '(123)456-7890')
        joy = ('Joy', '-123456---7890')
        marta = ('Marta', '123-456-7890')

        in_data = tablib.Dataset(headers=headers)
        in_data.append(john)
        in_data.append(george)
        in_data.append(binoy)
        in_data.append(jack)
        in_data.append(joy)
        in_data.append(marta)

        john = ('John', '')
        george = ('George', '')
        binoy = ('Binoy', '')
        jack = ('Jack', '123-456-7890')
        joy = ('Joy', '123-456-7890')
        marta = ('Marta', '123-456-7890')

        exp_data = tablib.Dataset(headers=headers)
        exp_data.append(john)
        exp_data.append(george)
        exp_data.append(binoy)
        exp_data.append(jack)
        exp_data.append(joy)
        exp_data.append(marta)

        self.perform_common_steps(test_command, in_data, exp_data)


    def test_validate_ssn(self):

        test_command = "do validate_ssn on 'SSN'"

        headers = ('Name', 'SSN')
        john = ('John', '')
        george = ('George', 'a')
        binoy = ('Binoy', '123')
        jack = ('Jack', '(123)456-789')
        joy = ('Joy', '-123456---789')
        marta = ('Marta', '123-45-6789')

        in_data = tablib.Dataset(headers=headers)
        in_data.append(john)
        in_data.append(george)
        in_data.append(binoy)
        in_data.append(jack)
        in_data.append(joy)
        in_data.append(marta)

        john = ('John', '')
        george = ('George', '')
        binoy = ('Binoy', '')
        jack = ('Jack', '123-45-6789')
        joy = ('Joy', '123-45-6789')
        marta = ('Marta', '123-45-6789')

        exp_data = tablib.Dataset(headers=headers)
        exp_data.append(john)
        exp_data.append(george)
        exp_data.append(binoy)
        exp_data.append(jack)
        exp_data.append(joy)
        exp_data.append(marta)

        self.perform_common_steps(test_command, in_data, exp_data)


    def test_validate_number(self):

        test_command = "do validate_number on 'frequency'"

        headers = ('Name', 'frequency')
        john = ('John', 'a')

        in_data = tablib.Dataset(headers=headers)
        in_data.append(john)

        exp_data = tablib.Dataset(headers=headers)

        with self.assertRaises(SystemExit) as cm:
            self.perform_common_steps(test_command, in_data, exp_data)

        self.assertEqual(cm.exception.code, -1)


        john = ('John', '01s')

        in_data = tablib.Dataset(headers=headers)
        in_data.append(john)

        with self.assertRaises(SystemExit) as cm:
            self.perform_common_steps(test_command, in_data, exp_data)

        self.assertEqual(cm.exception.code, -1)

        john = ('John', '')

        in_data = tablib.Dataset(headers=headers)
        in_data.append(john)
        exp_data.append(john)

        self.perform_common_steps(test_command, in_data, exp_data)

        john = ('John', 123.0)

        in_data = tablib.Dataset(headers=headers)
        in_data.append(john)
        exp_data = tablib.Dataset(headers=headers)
        exp_data.append(john)

        self.perform_common_steps(test_command, in_data, exp_data)


if __name__ == '__main__':
    unittest.main()