#!/usr/bin/python
# -*- coding: utf-8 -*-
from __future__ import absolute_import, division, generators, print_function, unicode_literals

import unittest
import os
from shutil import copyfile
import datetime

from ExcelProcessLibrary.excelmodule.keywords.definition.basic import BasicKeywords
from ExcelProcessLibrary.excelmodule.keywords.exception import *
from ExcelProcessLibrary import ExcelProcessLibrary


class TestBasicKeywords(unittest.TestCase):
	def __init__(self, *args, **kwargs):
		super(TestBasicKeywords, self).__init__(*args, **kwargs)
		self.excelLib = ExcelProcessLibrary()
		self.basic = BasicKeywords()

	@classmethod
	def setUpClass(cls):
		test_dir = os.path.dirname(__file__)
		test_file = os.path.join(test_dir, 'test_files', 'test_wb.xlsx')
		new_file = os.path.join(test_dir, 'test_files', 'test_wb_2.xlsx')
		copyfile(test_file, new_file)

	def setUp(self):
		self.excelLib.close_all_workbooks()

	def test_open_existing_workbook(self):
		"""Open Workbook"""
		self.excelLib.open_workbook("test_files/test_wb.xlsx", "one")
		num_of_wbs = len(self.basic._books)
		self.assertEqual(num_of_wbs, 1)

	def test_open_non_existing_workbook(self):
		"""Open Workbook"""
		with self.assertRaises(WorkbookNotFoundException):
			self.excelLib.open_workbook("test_wb.xlsx", "one")

	def test_open_already_opened_workbook(self):
		"""Open Workbook"""
		self.excelLib.open_workbook("test_files/test_wb.xlsx", "one")
		self.excelLib.open_workbook("test_files/test_wb.xlsx", "two")
		#self.basic.open_workbook("sam.xlsx", "three")
		num_of_wbs = len(self.basic._books)
		self.assertEqual(num_of_wbs, 2)

	def test_open_different_worbooks(self):
		"""Open Workbook"""
		self.excelLib.open_workbook("test_files/test_wb.xlsx", "one")
		self.excelLib.open_workbook("test_files/test_wb_2.xlsx", "another")
		#self.basic.open_workbook("sam.xlsx", "three")
		num_of_wbs = len(self.basic._books)
		self.assertEqual(num_of_wbs, 2)

	def test_close_all_workbooks(self):
		"""Close All Workbooks"""
		self.excelLib.open_workbook("test_files/test_wb.xlsx", "one")
		self.excelLib.close_all_workbooks()
		num_of_wbs = len(self.basic._books)
		self.assertEqual(num_of_wbs, 0)

	def test_close_all_workbooks_2(self):
		"""Close All Workbooks"""
		self.excelLib.close_all_workbooks()
		self.excelLib.close_all_workbooks()
		num_of_wbs = len(self.basic._books)
		self.assertEqual(num_of_wbs, 0)

	def test_close_existing_workbook_alias(self):
		"""Close Workbook"""
		self.excelLib.open_workbook("test_files/test_wb.xlsx", "one")
		self.excelLib.close_workbook("one")
		num_of_wbs = len(self.basic._books)
		self.assertEqual(num_of_wbs, 0)

	def test_close_existing_workbook_default(self):
		"""Close Workbook"""
		self.excelLib.open_workbook("test_files/test_wb.xlsx")
		self.excelLib.close_workbook()
		num_of_wbs = len(self.basic._books)
		self.assertEqual(num_of_wbs, 0)

	def test_close_not_opened_workbook(self):
		"""Close Workbook"""
		with self.assertRaises(WorkbookNotFoundException):
			self.excelLib.close_workbook("one")

	def test_save_workbook_as(self):
		"""Save Workbook As"""
		self.excelLib.open_workbook("test_files/test_wb.xlsx", "one")
		self.excelLib.save_workbook_as("test_files/test_wb_new.xlsx", "one")
		self.excelLib.close_workbook("one")
		num_of_wbs = len(self.basic._books)
		self.assertEqual(num_of_wbs, 0)
		test_dir = os.path.dirname(__file__)
		is_file = os.path.isfile(os.path.join(test_dir, 'test_files', 'test_wb_new.xlsx'))
		self.assertTrue(is_file)

	def test_save_as_wrong_path(self):
		"""Save Workbook As"""
		self.excelLib.open_workbook("test_files/test_wb.xlsx", "one")
		with self.assertRaises(DirectoryNotFoundException):
			self.excelLib.save_workbook_as("folder//sam_new_2", "one")
			self.excelLib.close_workbook("one")

	def test_save_as_full_path(self):
		"""Save Workbook As"""
		self.excelLib.open_workbook("test_files/test_wb.xlsx")
		new_file = os.path.join(os.getcwd(), str("ExcelProcessLibrary/tests/test_files/test_wb_new_2.xlsx"))
		self.excelLib.save_workbook_as(new_file)
		test_dir = os.path.dirname(__file__)
		is_file = os.path.isfile(os.path.join(test_dir, 'test_files', 'test_wb_new_2.xlsx'))
		self.assertTrue(is_file)

	def test_no_active_sheet(self):
		"""Set Active Worksheet"""
		self.excelLib.open_workbook("test_files/test_wb.xlsx")
		#print(self.basic._books["default_"])
		self.assertIsNone(self.basic._books["default_"].active_sheet)

	def test_activate_non_existing_sheet(self):
		"""Set Active Worksheet"""
		self.excelLib.open_workbook("test_files/test_wb.xlsx")
		with self.assertRaises(WorksheetNotFoundException):
			self.excelLib.set_active_worksheet("Sheet2")

	def test_activate_sheet_closed_workbook(self):
		"""Set Active Worksheet"""
		with self.assertRaises(WorkbookNotFoundException):
			self.excelLib.set_active_worksheet("Sheet2")

	def test_activate_existing_sheet(self):
		"""Set Active Worksheet"""
		self.excelLib.open_workbook("test_files/test_wb.xlsx")
		self.excelLib.set_active_worksheet("Sheet1")
		self.assertEqual(self.basic._books["default_"].active_sheet, "Sheet1")

	def test_activate_two_sheets(self):
		"""Set Active Worksheet"""
		self.excelLib.open_workbook("test_files/test_wb.xlsx", "one")
		self.excelLib.open_workbook("test_files/test_wb_2.xlsx", "two")
		self.excelLib.set_active_worksheet("Sheet1", "one")
		self.excelLib.set_active_worksheet("Sheet1", "two")
		self.assertEqual(self.basic._books["one"].active_sheet, "Sheet1")
		self.assertEqual(self.basic._books["two"].active_sheet, "Sheet1")

	def test_activate_two_sheets_2(self):
		"""Set Active Worksheet"""
		self.excelLib.open_workbook("test_files/test_wb.xlsx", "one")
		self.excelLib.open_workbook("test_files/test_wb_2.xlsx", "two")
		self.excelLib.set_active_worksheet("Sheet1", "one")
		self.excelLib.set_active_worksheet("Sheet1", "two")
		value_one = self.excelLib.get_cell_value("A1", workbook="one") # test_value
		self.excelLib.set_cell_value("A1", "test_value_two", workbook="two")
		value_two = self.excelLib.get_cell_value("A1", workbook="two") # test_value_two
		self.assertEqual(value_one, "test_value")
		self.assertEqual(value_two, "test_value_two")

	def test_new_worksheet(self):
		"""Create Worksheet"""
		self.excelLib.open_workbook("test_files/test_wb_2.xlsx")
		num_of_sheets = len(self.basic._books["default_"].wb_book.sheetnames)
		new_sheet = "new_sheet_0"
		self.excelLib.create_worksheet(new_sheet)
		num_of_sheets_after = len(self.basic._books["default_"].wb_book.sheetnames)
		self.assertEqual(num_of_sheets, num_of_sheets_after - 1)

	def test_new_worksheet_already_exist(self):
		"""Create Worksheet"""
		self.excelLib.open_workbook("test_files/test_wb_2.xlsx")
		new_sheet = "new_sheet_1"
		self.excelLib.create_worksheet(new_sheet)
		with self.assertRaises(Exception):
			self.excelLib.create_worksheet(new_sheet)

	def test_rename_worksheet(self):
		"""Rename Worksheet"""
		self.excelLib.open_workbook("test_files/test_wb_2.xlsx")
		new_name = "x_Sheet1"
		self.excelLib.rename_worksheet("Sheet1", new_name)
		exist = new_name in self.basic._books["default_"].wb_book.sheetnames
		self.assertTrue(exist)

	def test_activate_renamed_worksheet(self):
		"""Set Active Worksheet"""
		self.excelLib.open_workbook("test_files/test_wb_2.xlsx")
		self.excelLib.create_worksheet("a")
		self.excelLib.rename_worksheet("a", "x_Sheet2")
		self.excelLib.set_active_worksheet("x_Sheet2")
		self.assertEqual(self.basic._books["default_"].active_sheet, "x_Sheet2")

	def test_delete_existing_worksheet(self):
		"""Delete Worksheet"""
		self.excelLib.open_workbook("test_files/test_wb_2.xlsx")
		new_sheet = "3_new_sheet"
		self.excelLib.create_worksheet(new_sheet)
		self.excelLib.set_active_worksheet(new_sheet)
		self.excelLib.delete_worksheet(new_sheet)
		self.assertNotEqual(self.basic._books["default_"].active_sheet, "3_new_sheet")

	def test_delete_not_existing_worksheet(self):
		"""Delete Worksheet"""
		self.excelLib.open_workbook("test_files/test_wb_2.xlsx")
		with self.assertRaises(WorksheetNotFoundException):
			self.excelLib.delete_worksheet("4_new_sheet")

	def test_open_existing_workbook_absolute(self):
		"""Open Workbook"""
		test_dir = os.path.dirname(__file__)
		test_files_dir = os.path.join(test_dir, 'test_files')
		self.excelLib.open_workbook(os.path.join(test_files_dir, 'test_wb.xlsx'), "one")
		num_of_wbs = len(self.basic._books)
		self.assertEqual(num_of_wbs, 1)

	def test_open_existing_workbook_relative(self):
		"""Open Workbook"""
		self.excelLib.open_workbook("..//tests//test_files//test_wb.xlsx", "one")
		num_of_wbs = len(self.basic._books)
		self.assertEqual(num_of_wbs, 1)

	def test_open_existing_workbook_non_existing_relative(self):
		"""Open Workbook"""
		with self.assertRaises(WorkbookNotFoundException):
			self.excelLib.open_workbook("//nofolder//sam.xlsx", "one")
		num_of_wbs = len(self.basic._books)
		self.assertEqual(num_of_wbs, 0)

	def tearDown(self):
		self.excelLib.close_all_workbooks()

	@classmethod
	def tearDownClass(cls):
		test_dir = os.path.dirname(__file__)
		test_files_dir = os.path.join(test_dir, 'test_files')
		files = os.listdir(test_files_dir)
		for f in files:
			if f != "test_wb.xlsx":
				os.remove(os.path.join(test_files_dir, f))


class TestDataKeywords(unittest.TestCase):

	def __init__(self, *args, **kwargs):
		super(TestDataKeywords, self).__init__(*args, **kwargs)
		self.excelLib = ExcelProcessLibrary()
	#self.reader = DataReaderKeywords()
	#self.basic = BasicKeywords()

	def setUp(self):
		self.excelLib.close_all_workbooks()

	def test_get_cell_value(self):
		"""Get Cell Value"""
		self.excelLib.open_workbook("test_files/test_wb.xlsx")
		self.excelLib.set_active_worksheet("Sheet1")
		value = self.excelLib.get_cell_value("A1")
		self.assertEqual(value, "test_value")

	def test_get_cell_value_empty_cell(self):
		"""Get Cell Value"""
		self.excelLib.open_workbook("test_files/test_wb.xlsx")
		self.excelLib.set_active_worksheet("Sheet1")
		value = self.excelLib.get_cell_value("B1")
		self.assertIsNone(value)

	def test_get_cell_value_formula(self):
		"""Get Cell Value"""
		self.excelLib.open_workbook("test_files/test_wb.xlsx", "one")
		value = self.excelLib.get_cell_value("A2", "Sheet1", "one")
		self.assertEqual(value, 1)

	def test_get_cell_value_formula_active_sheet(self):
		"""Get Cell Value"""
		self.excelLib.open_workbook("test_files/test_wb.xlsx")
		self.excelLib.set_active_worksheet("Sheet1")
		value = self.excelLib.get_cell_value("A2")
		self.assertEqual(value, 1)

	def test_get_first_empty_cell_right_to_other(self):
		"""Get First Empty Cell Right To"""
		self.excelLib.open_workbook("test_files/test_wb.xlsx")
		self.excelLib.set_active_worksheet("Sheet1")
		value = self.excelLib.get_first_empty_cell_right_to("A1")
		self.assertEqual(value, "B1")

	def test_get_first_empty_cell_right_to_self(self):
		"""Get First Empty Cell Right To"""
		self.excelLib.open_workbook("test_files/test_wb.xlsx")
		self.excelLib.set_active_worksheet("Sheet1")
		value = self.excelLib.get_first_empty_cell_right_to("B1")
		self.assertEqual(value, "B1")

	def test_get_first_empty_cell_right_to_no_active_sheet(self):
		"""Get First Empty Cell Right To"""
		self.excelLib.open_workbook("test_files/test_wb.xlsx", "one")
		value = self.excelLib.get_first_empty_cell_right_to("B1", "Sheet1", "one")
		self.assertEqual(value, "B1")

	def test_get_first_empty_cell_below_other(self):
		"""Get First Empty Cell Below"""
		self.excelLib.open_workbook("test_files/test_wb.xlsx")
		self.excelLib.set_active_worksheet("Sheet1")
		value = self.excelLib.get_first_empty_cell_below("A1")
		self.assertEqual(value, "A3")

	def test_get_first_empty_cell_below_self(self):
		"""Get First Empty Cell Below"""
		self.excelLib.open_workbook("test_files/test_wb.xlsx")
		self.excelLib.set_active_worksheet("Sheet1")
		value = self.excelLib.get_first_empty_cell_below("A3")
		self.assertEqual(value, "A3")

	def test_get_first_empty_cell_below_no_active_sheet(self):
		"""Get First Empty Cell Below"""
		self.excelLib.open_workbook("test_files/test_wb.xlsx")
		value = self.excelLib.get_first_empty_cell_below("A1", "Sheet1")
		self.assertEqual(value, "A3")

	def test_get_cell_right_to(self):
		"""Get Cell Right To"""
		cell_ref = self.excelLib.get_cell_right_to("A1")
		self.assertEqual(cell_ref, "B1")

	def test_get_cell_left_to(self):
		"""Get Cell Left To"""
		cell_ref = self.excelLib.get_cell_left_to("B1")
		self.assertEqual(cell_ref, "A1")

	def test_get_cell_left_to_col_A(self):
		"""Get Cell Left To"""
		with self.assertRaises(CellNotFoundException):
			self.excelLib.get_cell_left_to("A1")

	def test_get_cell_below(self):
		"""Get Cell Below"""
		cell_ref = self.excelLib.get_cell_below("A1")
		self.assertEqual(cell_ref, "A2")

	def test_get_cell_above(self):
		"""Get Cell Above"""
		cell_ref = self.excelLib.get_cell_above("A2")
		self.assertEqual(cell_ref, "A1")

	def test_get_cell_above_row_1(self):
		"""Get Cell Above"""
		with self.assertRaises(CellNotFoundException):
			self.excelLib.get_cell_above("A1")

	def tearDown(self):
		self.excelLib.close_all_workbooks()


class TestEditKeywords(unittest.TestCase):

	def __init__(self, *args, **kwargs):
		super(TestEditKeywords, self).__init__(*args, **kwargs)
		self.excelLib = ExcelProcessLibrary()
	#self.reader = DataReaderKeywords()
	#self.basic = BasicKeywords()
	#self.edit = EditKeywords()

	@classmethod
	def setUpClass(cls):
		test_dir = os.path.dirname(__file__)
		test_file = os.path.join(test_dir, 'test_files', 'test_wb.xlsx')
		new_file = os.path.join(test_dir, 'test_files', 'test_wb_2.xlsx')
		copyfile(test_file, new_file)

	def setUp(self):
		self.excelLib.close_all_workbooks()

	def test_clear_cell(self):
		"""Clear Cell"""
		self.excelLib.open_workbook("test_files/test_wb_2.xlsx")
		self.excelLib.clear_cell("D1", "Sheet1")
		value = self.excelLib.get_cell_value("D1", "Sheet1")
		self.assertIsNone(value)

	def test_clear_cell_active_sheet(self):
		"""Clear Cell"""
		self.excelLib.open_workbook("test_files/test_wb_2.xlsx")
		self.excelLib.set_active_worksheet("Sheet1")
		self.excelLib.clear_cell("D1")
		value = self.excelLib.get_cell_value("D1")
		self.assertIsNone(value)

	def test_set_cell_value(self):
		"""Set Cell Value"""
		self.excelLib.open_workbook("test_files/test_wb_2.xlsx")
		self.excelLib.clear_cell("D1", "Sheet1")
		self.excelLib.set_cell_value("D1", "new value", "Sheet1")
		value = self.excelLib.get_cell_value("D1", "Sheet1")
		self.assertEqual(value, "new value")

	def test_set_cell_value_active_sheet(self):
		"""Set Cell Value"""
		self.excelLib.open_workbook("test_files/test_wb_2.xlsx")
		self.excelLib.set_active_worksheet("Sheet1")
		self.excelLib.clear_cell("D1")
		self.excelLib.set_cell_value("D1", "new value 2")
		value = self.excelLib.get_cell_value("D1")
		self.assertEqual(value, "new value 2")

	def test_add_formula_to_cell(self):
		"""Set Cell Value"""
		self.excelLib.open_workbook("test_files/test_wb_2.xlsx")
		self.excelLib.clear_cell("D2", "Sheet1")
		self.excelLib.set_cell_value("D2", "=SUM(A2:B2)", "Sheet1")
		value = self.excelLib.get_cell_value("D2", "Sheet1")
		self.excelLib.clear_cell("D2", "Sheet1")
		self.assertEqual(value, 1)

	def test_add_formula_to_cell_active_sheet(self):
		"""Set Cell Value"""
		self.excelLib.open_workbook("test_files/test_wb_2.xlsx")
		self.excelLib.set_active_worksheet("Sheet1")
		self.excelLib.clear_cell("D2")
		self.excelLib.set_cell_value("D2", "=SUM(A2:B2)")
		value = self.excelLib.get_cell_value("D2")
		self.excelLib.clear_cell("D2")
		self.assertEqual(value, 1)

	def test_add_list_to_row(self):
		"""Add List To Row"""
		self.excelLib.open_workbook("test_files/test_wb_2.xlsx")
		self.excelLib.set_active_worksheet("Sheet1")
		input_list = ['a', 'b', "c"]
		self.excelLib.add_list_to_row("D3", input_list)
		value_d = self.excelLib.get_cell_value("D3")
		value_e = self.excelLib.get_cell_value("E3")
		value_f = self.excelLib.get_cell_value("F3")
		self.excelLib.clear_cell("D3")
		self.excelLib.clear_cell("E3")
		self.excelLib.clear_cell("F3")
		self.assertListEqual(input_list, [value_d, value_e, value_f])

	def test_add_list_to_column(self):
		"""Add list To Column"""
		self.excelLib.open_workbook("test_files/test_wb_2.xlsx")
		self.excelLib.set_active_worksheet("Sheet1")
		input_list = ['a', 'b', "c"]
		self.excelLib.add_list_to_column("D3", input_list)
		value_3 = self.excelLib.get_cell_value("D3")
		value_4 = self.excelLib.get_cell_value("D4")
		value_5 = self.excelLib.get_cell_value("D5")
		self.excelLib.clear_cell("D3")
		self.excelLib.clear_cell("D4")
		self.excelLib.clear_cell("D5")
		self.assertListEqual(input_list, [value_3, value_4, value_5])

	def test_copy_cell_value(self):
		"""Copy Cell"""
		self.excelLib.open_workbook("test_files/test_wb_2.xlsx")
		self.excelLib.copy_cell("B6", "C6", "Sheet1")
		value_b = self.excelLib.get_cell_value("B6", "Sheet1")
		value_c = self.excelLib.get_cell_value("C6", "Sheet1")
		self.assertEqual(value_b, value_c)
		self.excelLib.clear_cell("C6", "Sheet1")

	def test_copy_cell_value_and_style(self):
		"""Copy Cell"""
		self.excelLib.open_workbook("test_files/test_wb_2.xlsx")
		self.excelLib.set_active_worksheet("Sheet1")
		self.excelLib.copy_cell("B7", "C7")
		value_b = self.excelLib.get_cell_value("B7")
		value_c = self.excelLib.get_cell_value("C7")
		self.assertEqual(value_b, value_c)
		self.excelLib.clear_cell("C7")

	def test_copy_cell_from_worksheet(self):
		"""Copy Cell From Worksheet"""
		self.excelLib.open_workbook("test_files/test_wb_2.xlsx")
		self.excelLib.create_worksheet("new_ws")
		self.excelLib.copy_cell_from_worksheet("A1", "A1", "Sheet1", "new_ws")
		value_original = self.excelLib.get_cell_value("A1", "Sheet1")
		value_new = self.excelLib.get_cell_value("A1", "new_ws")
		self.excelLib.delete_worksheet("new_ws")
		self.assertEqual(value_original, value_new)

	def test_set_cell_value_date(self):
		"""Set Cell Value"""
		self.excelLib.open_workbook("test_files/test_wb_2.xlsx")
		self.excelLib.set_active_worksheet("Sheet1")
		self.excelLib.set_cell_value("H2", "date:")
		self.excelLib.set_cell_value("I2", "1991.03.24")
		value_date = self.excelLib.get_cell_value("I2")
		self.assertEqual(value_date, datetime.datetime(1991, 3, 24, 0, 0))

	def test_set_cell_value_date_with_types(self):
		"""Set Cell Value"""
		self.excelLib.open_workbook("test_files/test_wb_2.xlsx")
		self.excelLib.set_active_worksheet("Sheet1")
		self.excelLib.set_cell_value("H3", "date:", cell_type="string")
		self.excelLib.set_cell_value("I3", "1991.03.24", cell_type="date")
		value_date = self.excelLib.get_cell_value("I2")
		self.assertEqual(value_date, datetime.datetime(1991, 3, 24, 0, 0))

	def test_set_cell_value_int(self):
		"""Set Cell Value"""
		self.excelLib.open_workbook("test_files/test_wb_2.xlsx")
		self.excelLib.set_active_worksheet("Sheet1")
		self.excelLib.set_cell_value("H4", "int:")
		self.excelLib.set_cell_value("I4", "27")
		value_int = self.excelLib.get_cell_value("I4")
		self.assertEqual(value_int, 27)

	def test_set_cell_value_int_with_types(self):
		"""Set Cell Value"""
		self.excelLib.open_workbook("test_files/test_wb_2.xlsx")
		self.excelLib.set_active_worksheet("Sheet1")
		self.excelLib.set_cell_value("H5", "int:", cell_type="string")
		self.excelLib.set_cell_value("I5", "27", cell_type="number")
		value_int = self.excelLib.get_cell_value("I5")
		self.assertEqual(value_int, 27)

	def test_set_cell_value_float(self):
		"""Set Cell Value"""
		self.excelLib.open_workbook("test_files/test_wb_2.xlsx")
		self.excelLib.set_active_worksheet("Sheet1")
		self.excelLib.set_cell_value("H6", "float:")
		self.excelLib.set_cell_value("I6", "3.14")
		value_float = self.excelLib.get_cell_value("I6")
		self.assertEqual(value_float, 3.14)

	def test_set_cell_value_float_with_types(self):
		"""Set Cell Value"""
		self.excelLib.open_workbook("test_files/test_wb_2.xlsx")
		self.excelLib.set_active_worksheet("Sheet1")
		self.excelLib.set_cell_value("H7", "float:", cell_type="string")
		self.excelLib.set_cell_value("I7", "3.14", cell_type="number")
		value_float = self.excelLib.get_cell_value("I7")
		self.assertEqual(value_float, 3.14)

	def tearDown(self):
		self.excelLib.close_all_workbooks()


class TestStyleKeywords(unittest.TestCase):

	def __init__(self, *args, **kwargs):
		super(TestStyleKeywords, self).__init__(*args, **kwargs)
		self.excelLib = ExcelProcessLibrary()

	#TODO egy minta sheetet kene letrehozni egy egy style_sheet kene a testsetupban
	# kezzel ossze lehetne haszonlitani a kettot

	@classmethod
	def setUpClass(cls):
		test_dir = os.path.dirname(__file__)
		test_file = os.path.join(test_dir, 'test_files', 'test_wb.xlsx')
		new_file = os.path.join(test_dir, 'test_files', 'test_wb_style.xlsx')
		copyfile(test_file, new_file)

	def setUp(self):
		self.excelLib.close_all_workbooks()

	def test_set_background_color_name(self):
		"""Set Cell Background"""
		self.excelLib.open_workbook("test_files/test_wb_style.xlsx")
		self.excelLib.set_cell_background("D2", "RED", "Sheet1")
		self.excelLib.clear_cell("D2", "Sheet1")
		self.excelLib.remove_style("D2", "Sheet1")
		self.excelLib.copy_cell("E2", "D2", "Sheet1")

	def test_set_background_color_code(self):
		"""Set Cell Background"""
		self.excelLib.open_workbook("test_files/test_wb_style.xlsx")
		self.excelLib.set_active_worksheet("Sheet1")
		self.excelLib.set_cell_value("D3", "orange background")
		self.excelLib.set_cell_background("D3", "00FFBB00")
		#self.excelLib.copy_cell("E3", "D3", "Sheet1")

	def test_set_font_color_name(self):
		"""Set Font Color"""
		self.excelLib.open_workbook("test_files/test_wb_style.xlsx")
		self.excelLib.set_active_worksheet("Sheet1")
		self.excelLib.set_cell_value("E3", "blue font color")
		self.excelLib.set_font_color("E3", "BLUE")
		#self.excelLib.copy_cell("E3", "D3", "Sheet1")
	#self.excelLib.clear_cell("D3")

	def test_set_font_color_code(self):
		"""Set Font Color"""
		self.excelLib.open_workbook("test_files/test_wb_style.xlsx")
		self.excelLib.set_cell_value("F3", "yellow font color", "Sheet1")
		self.excelLib.set_font_color("F3", "00FFBB00", "Sheet1")
		#self.excelLib.copy_cell("E3", "D3", "Sheet1")
	#self.edit.clear_cell("D3", "Sheet1")

	def test_set_cell_style(self):
		"""Set Cell Style"""
		self.excelLib.open_workbook("test_files/test_wb_style.xlsx")
		self.excelLib.set_cell_value("D4", "italic yellow underline", "Sheet1")
		self.excelLib.set_cell_style("D4", "italic", "Sheet1")
		self.excelLib.set_font_color("D4", "yellow", "Sheet1")
		self.excelLib.set_cell_style("D4", "underline", "Sheet1")
		#self.excelLib.copy_cell("E4", "D4", "Sheet1")

	def test_set_font_size(self):
		"""Set Font Size"""
		self.excelLib.open_workbook("test_files/test_wb_style.xlsx")
		self.excelLib.set_active_worksheet("Sheet1")
		self.excelLib.set_cell_value("E4", "24pt blue italic")
		self.excelLib.set_cell_style("E4", "italic")
		self.excelLib.set_font_color("E4", "blue")
		self.excelLib.set_font_size("E4", 24)
		#self.excelLib.copy_cell("E5", "D5")

	def test_remove_style(self):
		"""Remove Style"""
		self.excelLib.open_workbook("test_files/test_wb_style.xlsx")
		self.excelLib.set_cell_value("D6", "removed_style", "Sheet1")
		self.excelLib.set_cell_style("D6", "italic", "Sheet1")
		self.excelLib.set_font_color("D6", "blue", "Sheet1")
		self.excelLib.set_font_size("D6", 24, "Sheet1")
		self.excelLib.set_cell_value("D7", "removed_style", "Sheet1")
		self.excelLib.remove_style("D6", "Sheet1")
		self.excelLib.set_active_worksheet("Sheet1")
		self.excelLib.remove_style("D7")
		#self.excelLib.copy_cell("E6", "D6", "Sheet1")
		#elf.excelLib.copy_cell("E7", "D7", "Sheet1")

	def tearDown(self):
		self.excelLib.close_all_workbooks()


if __name__ == '__main__':
	unittest.main()
