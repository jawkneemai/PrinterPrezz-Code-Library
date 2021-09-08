# Johnathan Mai
# Ancillary Functions dealing with syntax, pathing, etc. to help with other functions

''' 
Directory of Functions

get_paths_from_folder(folder_path)
	Purpose: Gets paths to all files in a folder. 
	Input: string, path to folder
	Returns: list of strings, all of the paths in a list

make_folder(folder_name)
	Purpose: Makes a folder at current operating path with specified name.
	Input: folder_name(string)- desired name of folder.
	Returns: Nothing

get_file_name(file_path)
	Purpose: Parses the file name out of a path. No extension.
	Input: string, path to file
	Returns: string, solely the file name with no extension

'''

# Imports
import os

# Non-Python Imports


# Functions
def get_paths_from_folder(folder_path):
	paths = []
	for path in os.listdir(folder_path):
		paths.append(folder_path + '\\' + path)
	return paths

def make_folder(folder_name):
	folder_path = os.getcwd() + '\\' + folder_name
	if not os.path.isdir(folder_path):
		os.mkdir(folder_path)
	else: 
		print('Folder exists already')
	return folder_path

def get_file_name(file_path):
	try:
		slash_index = [pos for pos, char in enumerate(file_path) if char == '/']
		file_name = file_path[slash_index[len(slash_index)-1]+1:len(file_path)]
		file_name = file_name.split('.')[0]
	except:
		slash_index = [pos for pos, char in enumerate(file_path) if char == '\\']
		file_name = file_path[slash_index[len(slash_index)-1]+1:len(file_path)]
		file_name = file_name.split('.')[0]
	return file_name

def main():
	return

if __name__=='__main__':
	main()

print('Imported ancillary.py!')

