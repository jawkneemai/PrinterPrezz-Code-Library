# pull things from oqton?

import requests
from bs4 import BeautifulSoup
import csv 

# Gets the 100 rows of player data on a search query page. 
def getPlayerRows(url):	
	try:
		page = requests.get(url)
		#print(page.text)
		print(BeautifulSoup(page.text, 'html.parser').prettify())
		#rush_soup = BeautifulSoup(rush_page.content, 'html.parser')
		#rush_tbody = rush_soup.find('tbody')
		#rush_tr_elements = rush_tbody.find_all('tr')
	except AttributeError:
		print('Error creating or navigating Soup object')

	return 

# Takes a BeautifulSoup tag (from getPlayerRow, should be 100 element list of tags) and returns an object with all player data 
def parsePlayerRow(single_player_row):
	player_data = {}

	# Iterates through child's descendants to get innermost content, aka the player stat data
	for child in single_player_row.children:
		temp_data = ''
		if len(list(child.descendants)) > 0: # If a player stat has any content at all
			if len(list(child.descendants)) > 1: 
				if len(list(child.descendants)[-1]) > 1: temp_data = list(child.descendants)[-1]
				else: temp_data = list(child.descendants)[-2] # sometimes the deepest descendant is empty for whatever reason, goes to second deepest descendant
			else: temp_data = list(child.descendants)[0]
		player_data.add({child['data-stat']: temp_data})
	return player_data

# Takes playerData (list with all player data) and writes to specified csv file
def writePlayerRow(player_data, csv_file):
	fields = [] # can probably combine this with for loop to practice that one weird python list constructor thing
	# Gets stat names from player_data
	for key in player_data: fields.append(key)

	# Writes to CSV, with field names
	with open(csv_file, 'a', newline='') as file:
		writer = csv.DictWriter(file, fieldnames=fields)
		writer.writerow(player_data)
	return


def main():
	url = 'https://printerprezz-val.eu2.oqton.ai/parts?organise=&filter=&date=ALL&sort=-authored&state=ALL'
	getPlayerRows(url)

if __name__ == '__main__':
	main()