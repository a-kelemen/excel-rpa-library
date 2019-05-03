from __future__ import absolute_import, division, generators, print_function, unicode_literals


class ExcelLibraryException(Exception):
	pass
	#ROBOT_SUPPRESS_NAME = True


class WorksheetNotFoundException(ExcelLibraryException):
	pass


class WorkbookNotFoundException(ExcelLibraryException):
	pass


class CellNotFoundException(ExcelLibraryException):
	pass


class DirectoryNotFoundException(ExcelLibraryException):
	pass
