# Johnathan Mai
# Function: Analyzes multiple parts lists

# Python Imports
import os
from tkinter import *
from tkinter import filedialog

# Non-Python Imports
from pandas import read_excel

# Functions
def combine_parts_lists(*args):
	if not args:
		print('Choose your folder')
		try:
			root = Tk()
			root.withdraw()
			filePath = filedialog.askdirectory()
			xls_paths = get_paths_from_folder(filePath)
		except:
			print('No File Selected.')	
	elif len(args) == 1 and type(args[0]) == str:	
		try:
			xls_paths = get_paths_from_folder(args[0])
		except Exception as e:
			print(e.args[-1])
	else:
		sys.exit('Pass the folder path or leave arguments empty.')

	dfs = read_excel(xls_paths[0], engine='openpyxl')
	print(type(dfs))
	print(dfs)

	# Currently: .xlsx giving openpyxl reading errors because it's being downloaded programatically, 
	# if I download the PL manually it works..

'''
	for path in xls_paths:
		print(path.split('\\')[-1])
		dfs = read_excel(path, engine='openpyxl')

		df = pandas.DataFrame(final_rows)
		df.columns = data_fields
		df.to_excel(writer, sheet_name='Printer Log', index=False)
		writer.save()
'''

def count_parts():
	print('count_parts')

def count_coupons():
	print('count_coupons')

def get_paths_from_folder(folderPath):
	# Function: Returns paths of all files [list] in a folder
	# Returns: ^
	paths = []
	for path in os.listdir(folderPath):
		paths.append(folderPath + '\\' + path)
	return paths

def main():
	print('main()')

if __name__ == '__main__':
	main()