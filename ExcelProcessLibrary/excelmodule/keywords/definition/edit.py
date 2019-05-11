#!/usr/bin/python
# -*- coding: utf-8 -*-
from __future__ import absolute_import, division, generators, print_function, unicode_literals
from ..base import Base
from copy import copy
import datetime
import re


class EditKeywords(Base):

	def set_cell_value(self, reference, value, selected_sheet, workbook, cell_type):
		col, row = Base._get_coordinates_from_reference(reference)
		book = self._get_workbook(workbook)
		sheet = self._select_sheet(book, selected_sheet)
		cell = sheet.cell(row=row, column=col)
		value = str(value)
		if cell_type == "":
			if re.match(r"[0-9]+\.[0-9]{2}\.[0-9]{2}", value):
				cell_type = "date"
			else:
				try:
					float(value.replace(",", "."))
				except ValueError:
					cell_type = "string"
				else:
					cell_type = "number"

		if cell_type not in ["string", "number", "date"]:
			raise ValueError(cell_type + " is not a valid type!")

		if cell_type == "string":

			cell.value = str(value)
		if cell_type == "number":
			try:
				value = int(value)
				cell.number_format = "0"
				cell.value = value
			except ValueError:
				value = float(value.replace(",", "."))
				cell.number_format = "#,##0.00"
				cell.value = value

		if cell_type == "date":
			value = datetime.datetime.strptime(value, "%Y.%m.%d")
			cell.number_format = "YYYY.MM.DD"
			cell.value = value

		if value == "":
			cell.value = None
		book.wb_book.save(book.wb_path)

	def add_formula_to_cell(self, reference, formula, selected_sheet, workbook):
		self.set_cell_value(reference, formula, selected_sheet, workbook, "string")

	def clear_cell(self, reference, selected_sheet, workbook):
		self.set_cell_value(reference, "", selected_sheet, workbook, "string")

	def copy_cell(self, from_reference, to_reference, selected_sheet, workbook):
		book = self._get_workbook(workbook)
		sheet = self._select_sheet(book, selected_sheet)

		from_col, from_row = Base._get_coordinates_from_reference(from_reference)
		to_col, to_row = Base._get_coordinates_from_reference(to_reference)

		from_cell = sheet.cell(row=from_row, column=from_col)
		new_cell = sheet.cell(row=to_row, column=to_col, value=from_cell.value)

		if from_cell.value is None:
			new_cell.value = None
		if from_cell.has_style:
			new_cell.font = copy(from_cell.font)
			new_cell.border = copy(from_cell.border)
			new_cell.fill = copy(from_cell.fill)
			new_cell.number_format = copy(from_cell.number_format)
			new_cell.protection = copy(from_cell.protection)
			new_cell.alignment = copy(from_cell.alignment)
		book.wb_book.save(book.wb_path)

	def copy_cell_from_worksheet(self, from_sheet, to_sheet, from_reference, to_reference, workbook):
		book = self._get_workbook(workbook)

		from_col, from_row = Base._get_coordinates_from_reference(from_reference)
		from_sheet = self._select_sheet(book, from_sheet)

		to_col, to_row = Base._get_coordinates_from_reference(to_reference)
		to_sheet = self._select_sheet(book, to_sheet)

		from_cell = from_sheet.cell(row=from_row, column=from_col)
		new_cell = to_sheet.cell(row=to_row, column=to_col, value=from_cell.value)

		if from_cell.value is None:
			new_cell.value = None
		if from_cell.has_style:
			new_cell.font = copy(from_cell.font)
			new_cell.border = copy(from_cell.border)
			new_cell.fill = copy(from_cell.fill)
			new_cell.number_format = copy(from_cell.number_format)
			new_cell.protection = copy(from_cell.protection)
			new_cell.alignment = copy(from_cell.alignment)
		book.wb_book.save(book.wb_path)

	def copy_cell_from_workbook(self):
		pass

	def add_list_to_row(self, reference, items, selected_sheet, workbook, cell_type):
		iter_reference = reference
		for cell_item in items:
			self.set_cell_value(iter_reference, cell_item, selected_sheet, workbook, cell_type)
			iter_reference = self._cell_right_to(iter_reference)

	def add_list_to_column(self, reference, items, selected_sheet, workbook, cell_type):
		iter_reference = reference
		for cell_item in items:
			self.set_cell_value(iter_reference, cell_item, selected_sheet, workbook, cell_type)
			iter_reference = self._cell_below(iter_reference)
