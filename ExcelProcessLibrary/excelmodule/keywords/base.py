
from __future__ import absolute_import, division, generators, print_function, unicode_literals
from .exception import *
import sys


import os
import re
import string


from openpyxl import load_workbook
from openpyxl.styles import colors
from openpyxl.utils.cell import get_column_letter
import xlwings as xw
from robot.libraries.BuiltIn import BuiltIn, RobotNotRunningError


class Workb(object):
	#TODO atnevezni az adattagokat
	def __init__(self, book):
		self.wb_path = ""
		self.wb_name = ""
		self.wb_book = book
		self.folder = ""
		self.active_sheet = None

	def set_path(self, path):
		#norm_path = os.path.normpath(path)
		self.wb_path = path
		#abs_path = os.path.abspath(os.path.dirname(path))
		#print abs_path
		#self.folder = glob.glob(abs_path)[0]
		#print "ISDIR:", os.path.isdir(self.folder)
		#filename = os.path.basename(norm_path)
		self.folder = os.path.dirname(path)

	def set_name(self, name):
		self.wb_name = name

	def set_active_sheet(self, sheet):
		self.active_sheet = sheet
		#print("Workb-ben ez lett az aktiv: " + self.active_sheet)


class Base(object):
	
	#def __init__(self):
	#	self.fotag = "fotag"
	
	#a = Child1()

	#book = ""
	_books = {}
	_default_folder = "C:/"
	#def __init__(self):
	#	self.book = ""
	#	self._books = {}
	#	self._default_folder = "C:/"

	#self.probookss.append(openpyxl.load_workbook("..//..//..//files//majom.xlsx"))

	def _get_path(self, path):
		if not os.path.exists(path):
			path = os.path.join(self._default_folder, path)
			if not os.path.exists(path):
				raise WorkbookNotFoundException(str("Workbook not found: ") + path)
		#print("mar jo a folder: " + path)
		return path

	@staticmethod
	def _get_workbook(workbook):
		try:
			workbook = Base._books[workbook]
		except Exception:
			raise WorkbookNotFoundException("Workbook not found: " + workbook)
		return workbook

	@staticmethod
	def _get_coordinates_from_reference(ref):
		try:
			column_letters = re.search('^[A-Z]*', ref).group(0)
			row = re.search('[0-9]*$', ref).group(0)
		except AttributeError:
			raise Exception("hibas formatum")  #nem biztos
		num = 0
		for char in column_letters:
			if char in string.ascii_letters:
				num = 1 + num * 26 + (ord(char.upper()) - ord('A'))
		return int(num), int(row)

	#	def _get_default_workbook(self):
	#		try:
	#			workbook = self.books["default_"]
	#		except Exception:
	#			raise Exception("sajat exception jon ide: nincs megnyitve egy workbook se")
	#		return workbook

	#	def _get_workbook_alias(self, alias_name):
	#		try:
	#			workbook = self.books[alias_name]
	#		except Exception:
	#			raise Exception("sajat exception jon ide: nincs ilyen aliasu")
	#		return workbook

	#def use_folder(self, path):
	#	if not os.path.exists(path):
	#		raise Exception("nem letezik use_folder kwban megadott path")
	#	self._default_folder = path
	#	print(self._default_folder)


	def _select_sheet(self, workbook, selected_sheet):
		if selected_sheet == "active_":
			active_sheet = workbook.active_sheet
			if active_sheet is not None:
				return workbook.wb_book[active_sheet]
			else:
				raise WorksheetNotFoundException("No active sheet! Set active or add in argument!")
		else:
			for name in workbook.wb_book.sheetnames:
				if name == selected_sheet:
					return workbook.wb_book[name]
		raise WorksheetNotFoundException("Sheet not found: " + selected_sheet)

	# def _get_xl_type_name(self, type_char):
	# 	#https://openpyxl.readthedocs.io/en/stable/_modules/openpyxl/cell/cell.html
	# 	excel_data_types = {'s':"TYPE_STRING", 'f':"TYPE_FORMULA", 'n':"TYPE_NUMERIC", 'b':"TYPE_BOOL",
	# 	                    'inlineStr':"TYPE_INLINE", 'e':"TYPE_ERROR", 'str':"TYPE_FORMULA_CACHE_STRING"}
	# 	return excel_data_types[type_char]
	#
	# def _get_xl_number_from_typename(self, type_name):
	# 	# http://xlrd.readthedocs.io/en/latest/api.html
	# 	#NINCS NULL TYPE KISZEDTEM
	# 	excel_data_types = {'s':"TYPE_STRING", 'f':"TYPE_FORMULA", 'n':"TYPE_NUMERIC", 'b':"TYPE_BOOL",
	# 	                    'inlineStr':"TYPE_INLINE", 'e':"TYPE_ERROR", 'str':"TYPE_FORMULA_CACHE_STRING"}
	# 	#excel_data_types = {0: "XL_CELL_EMPTY",
	# 	#					1: "XL_CELL_TEXT",
	# 	#					2: "XL_CELL_NUMBER",
	# 	#					3: "XL_CELL_DATE",
	# 	#					4: "XL_CELL_BOOLEAN",
	# 	#					5: "XL_CELL_ERROR",
	# 	#					6: "XL_CELL_BLANK"}
	# 	return excel_data_types.keys()[excel_data_types.values().index(type_name)]

	def _is_formula(self, value):
		if str(value).startswith("="):
			return True
		else:
			return False

	def _get_formula_value(self, workbook, worksheet, row, col):
		app = xw.App(visible=False)
		wb = xw.Book(workbook)
		sht = None
		try:
			sht = wb.sheets[worksheet.title]
		except Exception:
			raise WorksheetNotFoundException("Sheet not found: " + worksheet)
		ref = self._to_reference(row, col)
		inter_value = sht.range(ref).value
		wb.close()
		app.kill()
		return inter_value

	def set_cell_type(self):
		"""
		"""
		pass

	def get_cell_type(self):
		"""
		"""
		pass

	def _nth_parent_folder(self, file_source, n):
		if n > 0:
			return self._nth_parent_folder(os.path.dirname(file_source), n-1)
		else:
			return file_source

	def _get_source(self, file_name):
		try:
			file_name = file_name.lstrip("\\").lstrip("/")
		except UnicodeDecodeError:
			pass
		if file_name == str("default_"):
			return self._get_process_folder()
		else:
			file_full_path = os.path.abspath(file_name)
			if not Base._is_absolute_path(file_name):
				return os.path.join(self._get_process_folder(), str(os.path.normpath(file_name)))
			else:
				return file_full_path

	@staticmethod
	def _is_absolute_path(path):
		return os.path.normpath(path) == os.path.abspath(path)

	def _get_process_folder(self):
		#print("-----------------------")
		#print(os.path.dirname(os.path.abspath(__file__)))
		import sys
		try:
			process_dir = os.path.dirname(BuiltIn().get_variable_value("${SUITE SOURCE}"))
			return process_dir
		except RobotNotRunningError as e:
			#print(type(e).__name__)
			#print("-----------------------")
			#print(os.path.dirname(os.path.abspath(__file__)))
			base_dir = self._nth_parent_folder(__file__, 3)
			#test directory of the module
			#print(os.path.join(base_dir, str("test")))
			return os.path.join(base_dir, str("tests"))



	@staticmethod
	def _get_color(color):
		if hasattr(colors, color.upper()):
			return getattr(colors, color.upper(), False)
		#print "NONET RETURNOLTAM"
		return None




	@staticmethod
	def _is_cell_empty(sheet, row, col):
		cell_value = sheet.cell(row=row, column=col).value
		if cell_value is None:
			return True
		return False



	@staticmethod
	def _to_reference(row, column):
		column_letter = get_column_letter(column)
		return column_letter + str(row)

	#def _cell_right_to(self, selected_sheet, reference, workbook="default_"):
	def _cell_right_to(self, reference):
		"""
		a mellette levo cellat adja vissza
		"""
		col, row = self._get_coordinates_from_reference(reference)
		col += 1
		return self._to_reference(row, col)

	#for cell in sheet[reference[0]]:
	#	print cell.value

	def _cell_below(self, reference):
		"""
		az alatta levo cellat adja vissza
		"""
		col, row = self._get_coordinates_from_reference(reference)
		row += 1
		return self._to_reference(row, col)

	#for cell in sheet[reference[0]]:
	#	print cell.value

	# NEM KELL. MINDIG MEG KELL ADNI HOFGY MELYIK SHEETEN DOGLOZUNK
	# def active_sheet_is(self, sheet="first", workbook="default_"):


# OK --OTT VOLT A HIBA HOGY A _books[] -nak adtam egy .path-t ami hiba
# OK --kell egy sajat workbook class aminek van egy openpyxl.workbook, egy path es egy name adattagja
# OK --ugyanabba a fajlba mentsen a set value
#ok dupla alabaszasok megszuntetese - https://stackoverflow.com/questions/41761645/when-to-use-one-or-two-underscore-in-python
#ok ?? python 3 -ban kene az egeszet atirni
#ok @keyword annotacio mukodjon -- kw-ket kulon mappaba kene, mukodesuktol fuggoen kulon fajlba (mint SeleniumLibary)
#
