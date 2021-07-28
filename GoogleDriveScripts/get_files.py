# Johnathan Mai
# Function: Obtaining access to files on google drive to do whatever with them

# Python Imports
import os
import io
from datetime import datetime as dt
import math

# Non Python Imports
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.http import MediaIoBaseDownload

# Functions
# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/drive.readonly', 
			'https://www.googleapis.com/auth/spreadsheets.readonly']

def download_file_google_api(googleApiService, fileObject, downloadPath):
	# Function: Downloads the file object from Google Drive to download path.
	# Returns: Nothing.
	# fileObject is the google API requested item (not google doc/sheets of course)
	if 'google-apps' in fileObject['mimeType']:
			print('Google Doc/Sheets file!\n')	
			# Fill in solution for google sheets/doc file ~~~~~~~~~~~~~~~~~~~~~~~
	else:
		request = googleApiService.files().get_media(fileId=fileObject['id'])
		fh = io.FileIO(downloadPath, 'wb')
		downloader = MediaIoBaseDownload(fh, request)
		done = False
		print('Downloading: ' + fileObject['name'])
		while done is False:
			status, done = downloader.next_chunk()
			print('Download %d%%. \n' % int(status.progress() * 100))

def convert_google_time(time_string):
	# Function: Converts the the 'createdTime' attribute of google API object to python datetime object.
	# Returns: Python datetime Object
	date, time = time_string[:-1].split('T') #  There's a weird Z at the end of it..
	hour, minute, seconds =  time.split(':')
	seconds_rounded = str(math.floor(float(seconds)))
	time = hour + ':' + minute + ':' + seconds_rounded

	return dt.strptime((date + ' ' + time), '%Y-%m-%d %H:%M:%S')

def search_drive(service, drive_id, query, page_token):
	# Function: Searches the drive ID for the intended query
	# Returns: Results of that search
	return service.files().list(
				pageSize=1000, 
				fields="nextPageToken, files(id, name, mimeType, createdTime, parents)", 
				includeItemsFromAllDrives=1, 
				supportsAllDrives=1,
				orderBy='modifiedTime', 
				driveId=drive_id,
				corpora='drive',
				pageToken=page_token,
				q=query).execute()

def get_files():
	# Function: Gets whatever files from the Google Drive from a query string. 
	# Returns: Folder Path of where the files are downloaded.

	# Credentialing and starting up google drive API service
	creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
	if os.path.exists('token.json'):
		creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
	if not creds or not creds.valid:
		if creds and creds.expired and creds.refresh_token:
			creds.refresh(Request())
		else:
			flow = InstalledAppFlow.from_client_secrets_file(
				'credentials.json', SCOPES)
			creds = flow.run_local_server(port=0)
		# Save the credentials for the next run
		with open('token.json', 'w') as token:
			token.write(creds.to_json())

	drive_service = build('drive', 'v3', credentials=creds)

	### EDIT THIS PART IF LOOKING FOR ANOTHER FILE ###

    # Currently trying to get specifically log files from Manufacturing > Projects > Printing Projects> PrinterPrezz Build (RT)
	drive_id = '0AC-_NOteXDL4Uk9PVA' # Manufacturing Drive

	# Edit this query for any file in Google Drive
	query_log_folder = "trashed=false and name='Log Files' and mimeType='application/vnd.google-apps.folder'"
	query_deposition_logs = "trashed=false and name contains 'deposition' and name contains '.log'"
    
    ### ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ ###

    # Gets all log files folders
	log_folder_ids = []
	page_token = None
	while True:
		results = search_drive(drive_service, drive_id, query_log_folder, page_token)		
		for folder in results.get('files', []):
			log_folder_ids.append(folder['id'])
		page_token = results.get('nextPageToken', None)
		if page_token is None:
			break

	# Get all deposition log files and see if their parent ID exists in folder list^
	log_files = []
	while True:
		results = search_drive(drive_service, drive_id, query_deposition_logs, page_token)
		for file in results.get('files', []):
			if file['parents'][0] in log_folder_ids:
				log_files.append(file)
		page_token = results.get('nextPageToken', None)
		if page_token is None:
			break

	# Downloading log files
	logs_folder_path = os.getcwd() + '\\log_files'
	if not os.path.isdir(logs_folder_path):
		os.mkdir(logs_folder_path)
	
	for file in log_files:
		temp_path = logs_folder_path + '\\' + file['name']
		download_file_google_api(drive_service, file, temp_path)

	return logs_folder_path

	# Where I left off: someone left a log file as a google doc and you have to use file export on it instead of downloading to get its contents 

def get_parts_list():
	# Function: Gets whatever files from the Google Drive from a query string. 
	# Returns: Folder Path of where the files are downloaded.

	# Credentials and drive_service Building
	creds = None
	if os.path.exists('token.json'):
		creds = Credentials.from_authorized_user_file('token.json', SCOPES)
	if not creds or not creds.valid:
		if creds and creds.expired and creds.refresh_token:
			creds.refresh(Request())
		else:
			flow = InstalledAppFlow.from_client_secrets_file(
				'credentials.json', SCOPES)
			creds = flow.run_local_server(port=0)
		with open('token.json', 'w') as token:
			token.write(creds.to_json())
	drive_service = build('drive', 'v3', credentials=creds)
	sheets_service = build('sheets', 'v4', credentials=creds)

    # Getting Parts List Google Sheets files from Traveler documents in Manufacturing > Projects > Printing Projects > PrinterPrezz Build (RT) > ~
	drive_id = '0AC-_NOteXDL4Uk9PVA' # Manufacturing Drive
	build_rt_id = '1lNxvM7WIKcGfM-XR-nAbP6nT0NtYPC7T'

	query_travelers = "trashed=false and mimeType='application/vnd.google-apps.folder' and name contains 'T4'"
	query_run_traveler = "trashed=false and mimeType='application/vnd.google-apps.folder' and name contains 'Run Traveler'"
	query_parts_lists = "trashed=false and mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'"

    ### ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ ###

    # Gets all log files folders

	page_token = None
	traveler_folder_ids = [] # This is the whole Traveler folder ("T####")
	while True:
		results = search_drive(drive_service, drive_id, query_travelers, page_token)
		for folder in results.get('files', []):
			if folder['parents'][0] == build_rt_id: # If "T####" folder has parent Printing Projects (RT)
				traveler_folder_ids.append(folder['id']) 
		page_token = results.get('nextPageToken', None)
		if page_token is None:
			break

	run_traveler_folder_ids = [] # This is "Run Traveler" folder inside the T#### Folders
	while True:
		results = search_drive(drive_service, drive_id, query_run_traveler, page_token)
		for folder in results.get('files', []):
			if folder['parents'][0] in traveler_folder_ids: # If "Run Traveler" folder has parent "T####"
				run_traveler_folder_ids.append(folder['id'])
		page_token = results.get('nextPageToken', None)
		if page_token is None:
			break

	parts_lists = [] # The parts list files in the "Run Traveler" folders.
	while True:
		results = search_drive(drive_service, drive_id, query_parts_lists, page_token)
		for xlsx in results.get('files', []):
			if xlsx['parents'][0] in run_traveler_folder_ids:
				#print(xlsx['id'] + ': ' + xlsx['name'] + ', ' + xlsx['mimeType'])
				#print(xlsx['parents'])
				parts_lists.append(xlsx)
		page_token = results.get('nextPageToken', None)
		if page_token is None:
			break

	parts_folder_path = os.getcwd() + '\\parts_lists'
	if not os.path.isdir(parts_folder_path):
		os.mkdir(parts_folder_path)
	
	for file in parts_lists:
		temp_path = parts_folder_path + '\\' + file['name']
		print(file['mimeType'])
		download_file_google_api(drive_service, file, temp_path)


	'''
	sheet_id = '12YyOQkhb8BkgVuDe-xqOPBQxVbnUjzMQ'
	sample_range = 'Class Data!A2:E'
	sheet = sheets_service.spreadsheets()
	result = sheet.values().get(spreadsheetId=sheet_id, range=sample_range).execute()
	values = result.get('values', [])

	for row in values:
		print('%s, %s' % (row[0], row[4]))
	'''

def main(): 
	print('What do you want to do?')
	#get_parts_list()
	#get_files()

if __name__ == '__main__':
	main()

print('Imported get_files.py!')