# Getting all log files from all print jobs and making them .xls

import os

from Library import get_files
from Library import printer_log
from Library import parts_list
from Library import ancillary
from Library import surface_area_from_ct
from Library import surface_area_from_stl as sastl

# Test Walkthrough for all Functions

'''
get_files.get_log_files()
folder_path = os.getcwd() + '\\Data\\log_files'
paths = ancillary.get_paths_from_folder(folder_path)
for path in paths:
	printer_log.parse_log(path)


folder_path = os.getcwd() + '\\Data\\ParsedLogs'
master_path = os.getcwd() + '\\Data\\Log_O2_Data.xlsx'
#printer_log.compile_data_from_logs(folder_path, master_path)
paths = ancillary.get_paths_from_folder(folder_path)
for path in paths:
	printer_log.plot_log_xy(path, 'Accumulated Time (s)', 'O2', add_to_previous_plot=True)

#printer_log.plot_log_xy(paths[5], 'Accumulated Time (s)', 'O2', title='lol')
#printer_log.plot_log_xy(paths[12], 'Accumulated Time (s)', 'O2', x_label='YES', add_to_previous_plot=True)


#get_files.get_parts_list()
folder_path = os.getcwd() + '\\Data\\parts_lists'
parts_list.combine_parts_lists(folder_path)
parts_list.count_parts()
parts_list.count_coupons()
'''

#printer_log.main()

sastl.main()