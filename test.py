# Getting all log files from all print jobs and making them .xls

import os

from Library import get_files
from Library import printer_log
from Library import parts_list
from Library import ancillary

'''
get_files.get_log_files()
folder_path = os.getcwd() + '\\Data\\log_files'
paths = ancillary.get_paths_from_folder(folder_path)
for path in paths:
	printer_log.parse_log(path)
'''
folder_path = os.getcwd() + '\\Data\\log_files'
master_path = os.getcwd() + '\\Data\\Log_O2_Data'
printer_log.compile_data_from_logs(folder_path, master_path)

# Running through all code, currently compile_data_from_logs is giving bad zip file :/

'''
folderPath = os.getcwd() + '\\log_files'
folder_path = os.getcwd()+'\\O2Logs.xlsx'
print(parse_logs.get_file_name(folder_path))

#folderPath = get_files.get_files()

filePaths = parse_logs.get_paths_from_folder(folderPath)
for path in filePaths:
	if '.log' in path:
		parse_logs.parse_log(path)
		print('Parsed ' + parse_logs.get_file_name(path))
	else:
		print(path + '\n' + 'is not a log file.')
#folderPath = os.getcwd() + '\\ParsedLogs'
#parse_logs.compile_data_from_logs(folderPath, os.getcwd()+'\\O2Logs.xlsx')


#get_files.get_files()
#pl.combine_parts_lists(folderPath)
#pl.count_parts()
#pl.count_coupons()




'''