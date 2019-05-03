#!/usr/bin/python
# -*- coding: utf-8 -*-
from __future__ import absolute_import, division, generators, print_function, unicode_literals
from ..base import Base
from ..exception import *


class DataReaderKeywords(Base):

	def get_cell_value(self, reference, selected_sheet, workbook):
		col, row = Base._get_coordinates_from_reference(reference)
		book = self._get_workbook(workbook)
		sheet = self._select_sheet(book, selected_sheet)
		raw_value = sheet.cell(row=row, column=col).value
		if raw_value is None:
			return None
		else:
			if not self._is_formula(raw_value):
				return sheet.cell(row=row, column=col).value
			else:
				intern_value = self._get_formula_value(book.wb_path, sheet, row, col)
				return intern_value
			#TODO tipusos dolgol xlrd tipusosak kulon FAJL staticba

	def get_first_empty_cell_below(self, reference, selected_sheet, workbook): #below??
		col, row = Base._get_coordinates_from_reference(reference)
		book = self._get_workbook(workbook)
		sheet = self._select_sheet(book, selected_sheet)
		while not self._is_cell_empty(sheet, row, col):
			row += 1
		return self._to_reference(row, col)

	def get_first_empty_cell_right_to(self, reference, selected_sheet, workbook):
		# TODO next to
		col, row = Base._get_coordinates_from_reference(reference)
		book = self._get_workbook(workbook)
		sheet = self._select_sheet(book, selected_sheet)
		while not self._is_cell_empty(sheet, row, col):
			col += 1
		return self._to_reference(row, col)

	def get_cell_right_to(self, reference):
		# TODO Next to
		col, row = Base._get_coordinates_from_reference(reference)
		return self._to_reference(row, col + 1)

	def get_cell_left_to(self, reference):
		# TODO Next to
		col, row = Base._get_coordinates_from_reference(reference)
		if col > 1:
			return self._to_reference(row, col - 1)
		else:
			raise CellNotFoundException("There is no cell left to column A!")

	def get_cell_below(self, reference):
		# TODO Next to
		col, row = Base._get_coordinates_from_reference(reference)
		return self._to_reference(row + 1, col)

	def get_cell_above(self, reference):
		# TODO Next to
		col, row = Base._get_coordinates_from_reference(reference)
		if row > 1:
			return self._to_reference(row - 1, col)
		else:
			raise CellNotFoundException("There is no cell above row 1!")


