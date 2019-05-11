import setuptools


setuptools.setup(
	name='excelprocesslibrary',
	version='0.1',
	description='Library for excel automation.',
	url='',
	author='Andras Kelemen',
	author_email='kelemenandras11@gmail.com',
	license='MIT',
	packages=setuptools.find_packages(exclude=['ExcelProcessLibrary.tests']),

	install_requires=[
		'openpyxl',
		'robotframework',
		'xlwings'
	],
	test_suite="tests",
	tests_require=['nose'],
	zip_safe=False
)
