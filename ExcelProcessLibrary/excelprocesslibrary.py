#!/usr/bin/python
# -*- coding: utf-8 -*-
from __future__ import absolute_import, division, generators, print_function, unicode_literals

from .excelmodule.keywords.definition.edit import EditKeywords
from .excelmodule.keywords.definition.basic import BasicKeywords
from .excelmodule.keywords.definition.datareader import DataReaderKeywords
from .excelmodule.keywords.definition.style import StyleKeywords


def keyword(name=None):
	#"from robot.api.deco import keyword" alapjan
	#TODO ebbe az kell, hogy lefuttassa a keywordot, es ha sikres, akkor kiir valamit hogy sikeres hogy legyen
	#valami haszna
	if callable(name):
		return keyword()(name)

	def _method_wrapper(func):
		func.robot_name = name
		return func
	return _method_wrapper


#TODO tesztelni aq copy range tipusokkal
class ExcelProcessLibrary(object):
	__version__ = '0.1.0'

	ROBOT_LIBRARY_SCOPE = 'GLOBAL'
	ROBOT_LIBRARY_VERSION = __version__

	def __init__(self):
		self.edit = EditKeywords()
		self.style = StyleKeywords()
		self.basic = BasicKeywords()
		self.data_reader = DataReaderKeywords()

	@keyword(name="Open Workbook")
	def open_workbook(self, name, alias=None):
		"""
		Opens the given Excel file, defined by ``name``.

		If the optional ``alias`` argument is given, alias name could be used as a filename in tests.

		If an alias name is not defined, the default workbook is going to be the active one.

		Examples:
		| `Open Workbook` | ${CURDIR}\sample.xlsx    |        |    |       | # opens _sample.xlsx_ and sets it to default |
		| `Open Workbook` | ${CURDIR}\\another.xlsx  | xls_1  |    |       | # opens _another.xlsx_ and sets an alias     |
		| ${val_sample}=  | `Get Cell Value`         | sheet1 | A1 |       | # returns a value from _sample.xlsx_         |
		| ${val_another}= | `Get Cell Value`         | sheet1 | A1 | xls_1 | # returns a value from _another.xlsx_        |
		"""
		self.basic.open_workbook(name, alias)

	@keyword(name="Close Workbook")
	def close_workbook(self, workbook="default_"):
		"""
		Closes the given workbook.

		Workbook could be referenced by the alias name. If there is no argument, it closes the default workbook.

		Examples:
		| `Open Workbook`  | ${CURDIR}\sample.xlsx   |        | # opens _sample.xlsx_ and sets it to default |
		| `Open Workbook`  | ${CURDIR}\\another.xlsx | xls_1  | # opens _another.xlsx_ and sets an alias     |
		| `Close Workbook` | xls_1                   |        | # closes _another.xlsx_ by its alias name    |
		| `Close Workbook` |                         |        | # closes _sample.xlsx_ (default workbook)    |
		"""
		self.basic.close_workbook(workbook)

	@keyword(name="Close All Workbooks")
	def close_all_workbooks(self):
		"""
		Closes all opened workbooks.

		Examples:
		| `Open Workbook`       | ${CURDIR}\sample.xlsx   |       | # opens _sample.xlsx_ and sets it to default   |
		| `Open Workbook`       | ${CURDIR}\\another.xlsx | xls_1 | # opens _another.xlsx_ and sets an alias       |
		| `Close All Workbooks` |                         |       | # closes both _sample.xlsx_ and _another.xlsx_ |
		"""
		self.basic.close_all_workbooks()

	@keyword(name="Create Worksheet")
	def create_worksheet(self, new_sheet_name, workbook="default_"):
		"""
		Creates a new worksheet with name ``new_sheet_name`` to the given ``workbook``.

		Examples:
		| `Open Workbook`    | ${CURDIR}\sample.xlsx   |       | # opens _sample.xlsx_ and sets it to default                              |
		| `Open Workbook`    | ${CURDIR}\\another.xlsx | xls_1 | # opens _another.xlsx_ and sets an alias                                  |
		| `Create Worksheet` | new_sheet_01            | xls_1 | # creates a new sheet named _new_sheet_01_ to the _another.xlsx_ workbook |
		| `Create Worksheet` | new_sheet_01            |       | # creates a new sheet named _new_sheet_01_ to the _sample.xlsx_ workbook  |

		"""
		self.basic.create_worksheet(new_sheet_name, workbook)

	@keyword(name="Delete Worksheet")
	def delete_worksheet(self, sheet_name, workbook="default_"):
		"""
		Removes a worksheet with the given ``sheet_name`` from the given ``workbook``.

		Examples:
		| `Open Workbook`    | ${CURDIR}\sample.xlsx   |        | # opens _sample.xlsx_ and sets it to default workbook           |
		| `Open Workbook`    | ${CURDIR}\\another.xlsx | xls_1  | # opens _another.xlsx_ and sets an alias                        |
		| `Delete Worksheet` | new_sheet_01            | xls_1  | # deletes _new_sheet_01_ worksheet from _another.xlsx_ workbook |
		| `Delete Worksheet` | new_sheet_01            |        | # deletes _new_sheet_01 worksheet_ from _sample.xlsx_ workbook  |
		"""
		self.basic.delete_worksheet(sheet_name, workbook)

	@keyword(name="Set Active Worksheet")
	def set_active_worksheet(self, selected_sheet, workbook="default_"):
		"""
		Sets the active worksheet of the given ``workbook``.

		Examples:
		| `Open Workbook`        | ${CURDIR}\sample.xlsx  |         |                  | # opens _sample.xlsx_ and sets it to default workbook                |
		| `Set Active Worksheet` | Sheet1                 |         |                  | # sets _Sheet1_ worksheet to active                                  |
		| ${value}=              | `Get Cell Value`       | A1      |                  | # returns the value of A1 cell                                       |
		| `Open Workbook`        | another.xlsx           | another |                  |                                                                      |
		| `Set Active Worksheet` | Sheet1                 | another |                  | # sets _Sheet1_ worksheet to active in _another_wb_ workbook         |
		| ${value}=              | `Get Cell Value`       | A1      | workbook=another | # returns the value of A1 cell of _another_wb_ (kwargs must be used) |
		"""
		self.basic.set_active_worksheet(selected_sheet, workbook)

	@keyword(name="Rename Worksheet")
	def rename_worksheet(self, selected_sheet, new_sheet_name, workbook="default_"):
		"""
		Renames the given worksheet.

		``selected_sheet`` and ``new_sheet_name`` parameters are mandatory.

		Examples:
		| `Rename Worksheet` | human | robot  |       | # renames the worksheet named _human_ to _robot_ in default workbook |
		| `Rename Worksheet` | past  | future | alias | # renames the worksheet named _past_ to _future_ in _alias_ workbook |
		"""
		self.basic.rename_worksheet(selected_sheet, new_sheet_name, workbook)

	@keyword(name="Save Workbook As")
	def save_workbook_as(self, new_name, workbook="default_"):
		"""
		Saves the workbook as ``new_name``.

		Examples:
		| `Open Workbook`    | ${CURDIR}\sample.xlsx           | # opens _sample.xlsx_ and sets it to default                                   |
		| `Save Workbook As` | sample_copy.xlsx                | # saves the workbook as _sample_copy.xlsx_ to the directory of the .robot file |
		| `Save Workbook As` | ${downloads}\sample_copy_2.xlsx | # saves the workbook as _sample_copy_2.xlsx_ to the given directory            |
		"""
		self.basic.save_workbook_as(new_name, workbook)

	@keyword(name="Get Cell Value")
	def get_cell_value(self, reference, selected_sheet="active_", workbook="default_"):
		"""
		Returns the value of a cell of the ``selected_sheet``.
		``reference`` argument is an Excel reference to the given cell.

		If the optional ``workbook`` argument is not given, ``selected_sheet`` and cell is taken from the default workbook.

		If the cell contains a formula, this keyword returns the value of the formula.

		Examples:
		| `Open Workbook` | ${CURDIR}\sample.xlsx   |        |        |       |                                                                |
		| `Open Workbook` | ${CURDIR}\\another.xlsx | xls_1  |        |       |                                                                |
		| ${val}=         | `Get Cell Value`        | A1     | sheet1 |       | # returns a value from sample.xlsx, cell A1, sheet1 worksheet  |
		| ${val2}=        | `Get Cell Value`        | A1     | sheet1 | xls_1 | # returns a value from another.xlsx, cell A1, sheet1 worksheet |
		"""
		return self.data_reader.get_cell_value(reference, selected_sheet, workbook)

	@keyword(name="Get First Empty Cell Below")
	def get_first_empty_cell_below(self, reference, selected_sheet="active_", workbook="default_"):
		"""
		Returns the first empty cell below the given cell.
		If the cell is empty, the given ``reference`` will be returned.

		Examples:
		| ${cell}= | `Get First Empty Cell Below` | A1 | sheet1 | # returns A4 if neither A1, A2, nor A3 is empty, but A4 is empty |
		| ${cell}= | `Get First Empty Cell Below` | A1 | sheet1 | # returns A1 if A1 is empty                                      |

		Returns a reference to the cell.
		"""
		return self.data_reader.get_first_empty_cell_below(reference, selected_sheet, workbook)

	@keyword(name="Get First Empty Cell Right To")
	def get_first_empty_cell_right_to(self, reference, selected_sheet="active_", workbook="default_"):
		"""
		Returns the first empty cell right to the given cell.

		Examples:
		| ${cell}= | `Get First Empty Cell Right To` | A1 | sheet1 | # returns D1 if neither A1, B1, nor C1 is empty, but D1 is empty |
		| ${cell}= | `Get First Empty Cell Right To` | A1 | sheet1 | # returns A1 if A1 is empty                                      |

		Returns a reference to the cell.
		"""
		return self.data_reader.get_first_empty_cell_right_to(reference, selected_sheet, workbook)

	@keyword(name="Get Cell Right To")
	def get_cell_right_to(self, reference):
		"""
		Returns the cell reference right to the given cell.

		Examples:
		| ${cell_ref}= |`Get Cell Right To` | A1 | # returns B1 |
		"""
		return self.data_reader.get_cell_right_to(reference)

	@keyword(name="Get Cell Left To")
	def get_cell_left_to(self, reference):
		"""
		Returns the cell reference left to the given cell.

		Examples:
		| ${cell_ref}= |`Get Cell Left To` | B1 | # returns A1 |
		| ${cell_ref}= |`Get Cell Left To` | A1 | # fails      |
		"""
		return self.data_reader.get_cell_left_to(reference)

	@keyword(name="Get Cell Above")
	def get_cell_above(self, reference):
		"""
		Returns the cell reference above the given cell.

		Examples:
		| ${cell_ref}= |`Get Cell Above` | A2 | # returns A1 |
		| ${cell_ref}= |`Get Cell Above` | A1 | # fails      |
		"""
		return self.data_reader.get_cell_above(reference)

	@keyword(name="Get Cell Below")
	def get_cell_below(self, reference):
		"""
		Returns the cell reference below the given cell.

		Examples:
		| ${cell_ref}= |`Get Cell Below` | A1 | # returns A2 |
		"""
		return self.data_reader.get_cell_below(reference)

	@keyword(name="Set Cell Value")
	def set_cell_value(self, reference, value, selected_sheet="active_", workbook="default_", cell_type=""):
		"""
		Sets the ``value`` of the given cell of the ``selected_sheet``.
		``reference`` argument is an Excel reference to the given cell.

		If the optional ``workbook`` argument is not given, the default workbook will be active.
		If the ``cell_type`` argument is not given, keyword tries to convert the value to number or date. If it isn't possible, the cell type is going to be string.

		Possible ``cell_type`` values:
		- number
		- date
		- string

		Examples:
		| `Open Workbook`  | sample.xlsx  |            |        |       |                  |                                                                            |
		| `Open Workbook`  | another.xlsx | xls_1      |        |       |                  |                                                                            |
		| `Set Cell Value` | A1           | 14         | sheet1 |       |                  | # sets the value of A1 cell in _sheet1_ worksheet, _sample.xlsx_ (number)  |
		| `Set Cell Value` | A1           | spagetti   | sheet1 | xls_1 |                  | # sets the value of A1 cell in _sheet1_ worksheet, _another.xlsx_ (string) |
		| `Set Cell Value` | A2           | 1991.03.24 | sheet1 | xls_1 |                  | # sets the value of A2 cell in _sheet1_ worksheet, _another.xlsx_ (date)   |
		| `Set Cell Value` | A3           | 14         | sheet1 |       | cell_type=string | # value of A3 cell is going to be _14_ (string)                            |

		If you want to add a formula to the cell, use `Add Formula To Cell` instead.
		"""
		self.edit.set_cell_value(reference, value, selected_sheet, workbook, cell_type)

	@keyword(name="Add Formula To Cell")
	def add_formula_to_cell(self, reference, formula, selected_sheet="active_", workbook="default_"):
		"""
		Adds a formula to the given cell of the ``selected_sheet``.

		``reference`` argument is an Excel reference to the given cell.

		If the optional ``workbook`` argument is not given, the default workbook will be active.

		Examples:
		| `Open Workbook`        | sample.xlsx  |             |        |       |                                                                   |
		| `Open Workbook`        | another.xlsx | xls_1       |        |       |                                                                   |
		| `Add Formula To Cell`  | A6           | =SUM(A1:A5) | sheet1 |       | # adds a formula to A6 cell in _sheet1_ worksheet, _sample.xlsx_  |
		| `Add Formula To Cell`  | A6           | =SUM(A1:A5) | sheet1 | xls_1 | # adds a formula to A6 cell in _sheet1_ worksheet, _another.xlsx_ |

		If you want to add a simple value to the cell, use `Set Cell Value` instead.
		"""
		self.edit.add_formula_to_cell(reference, formula, selected_sheet, workbook)

	@keyword(name="Clear Cell")
	def clear_cell(self, reference, selected_sheet="active_", workbook="default_"):
		"""
		Removes the value from the given cell.

		Examples:
		| `Clear Cell` | A1 | sheet1 |       | # value of A1 cell in _sheet1_ (in default workbook) is going to be None       |
		| `Clear Cell` | A1 | sheet1 | alias | # value of A1 cell in _sheet1_ (in workbook named _alias_) is going to be None |
		"""
		self.edit.clear_cell(reference, selected_sheet, workbook)

	@keyword(name="Copy Cell")
	def copy_cell(self, from_reference, to_reference, selected_sheet="active_", workbook="default_"):
		"""
		Copies a cell in a worksheet.

		Examples:
		| `Copy Cell` | A1 | C1 | sheet1 | # copies A1 cell to C1 cell |

		If you want to copy from or to an other worksheet use ``Copy Cell From Worksheet`` instead.
		"""
		self.edit.copy_cell(from_reference, to_reference, selected_sheet, workbook)

	@keyword(name="Copy Cell From Worksheet")
	def copy_cell_from_worksheet(self, from_reference, to_reference, from_sheet, to_sheet, workbook="default_"):
		"""
		Copies a cell in a workbook.

		``from_reference``, ``to_reference``, ``from_sheet`` and ``to_sheet`` arguments are mandatory.

		Examples:
		| `Copy Cell From Worksheet` | A1 | C1 | sheet1 | sheet2 |     | # copies A1 cell to C1 cell from _sheet1_ sheet to _sheet2_ sheet in default workbook     |
		| `Copy Cell From Worksheet` | A1 | C1 | sheet1 | sheet2 | wb2 | # copies A1 cell to C1 cell from _sheet1_ sheet to _sheet2_ sheet in workbook named _wb2_ |
		"""
		self.edit.copy_cell_from_worksheet(from_sheet, to_sheet, from_reference, to_reference, workbook)

	@keyword(name="Add List To Row")
	def add_list_to_row(self, reference, items, selected_sheet="active_", workbook="default_", cell_type=""):
		"""
		Adds a list to the row, starting from the given cell.

		If the optional ``workbook`` argument is not given, the default workbook will be active.
		If the ``cell_type`` argument is not given, keyword tries to convert the value to number or date. If it isn't possible, the cell type is going to be string.

		Examples:
		| `Add List To Row` | H1 | ${list} | sheet1 | workbook1 | # adds the _${list}_ to the _H_ row starting from H1 cell |
		"""

		self.edit.add_list_to_row(reference, items, selected_sheet, workbook, cell_type)

	@keyword(name="Add List To Column")
	def add_list_to_column(self, reference, items, selected_sheet="active_", workbook="default_", cell_type=""):
		"""
		Adds a list to the column, starting from the given cell.

		If the optional ``workbook`` argument is not given, the default workbook will be active.
		If the ``cell_type`` argument is not given, keyword tries to convert the value to number or date. If it isn't possible, the cell type is going to be string.

		Examples:
		| `Add List To Column` | H1 | ${list} | sheet1 | workbook1 | # adds the _${list}_ to the _H_ column starting from H1 cell |
		"""
		self.edit.add_list_to_column(reference, items, selected_sheet, workbook, cell_type)

	@keyword(name="Set Cell Background")
	def set_cell_background(self, reference, color, selected_sheet="active_", workbook="default_"):
		"""
		Sets the background of the cell.

		``color`` is RGB or aRGB hexvalue.

		Examples:
		| `Set Cell Background` | A1 | YELLOW   | sheet2 | # background of A1 is going to be yellow in the default workbook |
		| `Set Cell Background` | A2 | 00FFBB00 | sheet2 | # sets the background of A2 cell                                 |
		"""
		self.style.set_cell_background(reference, color, selected_sheet, workbook)

	@keyword(name="Set Font Color")
	def set_font_color(self, reference, font_color, selected_sheet="active_", workbook="default_"):
		"""
		Sets the color of the font in the given cell.

		``font_color`` is RGB or aRGB hexvalue.

		Examples:
		| `Set Font Color` | A1 | 00FFBB00 | sheet2 | # sets the font color of A2 cell |
		"""
		self.style.set_font_color(reference, font_color, selected_sheet, workbook)

	@keyword(name="Set Cell Style")
	def set_cell_style(self, reference, style, selected_sheet="active_", workbook="default_"):
		"""
		Sets the style of the cell.

		Possible ``style`` values:
		- italic
		- bold
		- underline
		- normal (removes style)

		Examples:
		| `Set Cell Style` | A1 | bold   | sheet1 |     | # style of A1 cell is going to be bold in _sheet1_ worksheet of the default workbook |
		| `Set Cell Style` | A2 | italic | sheet1 | wb1 | # style of A2 cell is going to be italic in _sheet1_ worksheet of _wb1_ workbook     |
		"""
		self.style.set_cell_style(reference, style, selected_sheet, workbook)

	@keyword(name="Set Font Size")
	def set_font_size(self, reference, font_size, selected_sheet="active_", workbook="default_"):
		"""
		Sets the font size.

		Examples:
		| `Set Font Size` | A1 | 12 | sheet1                |     | # font size of A1 cell is going to be 12 pts in _sheet1_ worksheet of the default workbook |
		| `Set Font Size` | A2 | 20 | selected_sheet=sheet1 | wb1 | # font size of A2 cell is going to be 20 pts in _sheet1_ worksheet of _wb1_ workbook       |
		"""
		self.style.set_font_size(reference, font_size, selected_sheet, workbook)

	@keyword(name="Remove Style")
	def remove_style(self, reference, selected_sheet="active_", workbook="default_"):
		"""
		Removes the style of the given cell.

		Examples:
		| `Remove Style` | A1 | sheet1 | # removes the style of A1 cell in the _sheet1_ worksheet of the default workbook |
		"""
		self.style.remove_style(reference, selected_sheet, workbook)
