# Getting all log files from all print jobs and making them .xls

import os

from GoogleDriveScripts import get_files
from PrinterLogScripts import parse_logs
from ManufacturingScripts import analyze_parts_list as pl


folderPath = os.getcwd() + '\\log_files'
#folderPath = get_files.get_files()
'''
filePaths = parse_logs.get_paths_from_folder(folderPath)
for path in filePaths:
	if '.log' in path:
		parse_logs.parse_log(path)
		print('Parsed ' + parse_logs.get_file_name(path))
	else:
		print(path + '\n' + 'is not a log file.')
'''
folderPath = os.getcwd() + '\\ParsedLogs'
parse_logs.compile_data_from_logs(folderPath, os.getcwd()+'\\O2Logs.xlsx')


#get_files.get_files()
#pl.combine_parts_lists(folderPath)
#pl.count_parts()
#pl.count_coupons()




