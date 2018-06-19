import json
import requests
import csv
import secrets as secrets
from bs4 import BeautifulSoup
import pandas as pd
import re

from openpyxl import *

from openpyxl import load_workbook

import openpyxl

# api_key = secrets.api_key

FNAME = "Judeo_arabic_spreadsheet.xlsx"



class Document():
	def __init__(self, title):
		self.title = title
		self.subject_tags = ""

	def __str__(self):
		return subject_tags
	pass

def write_to_excel(filename, current_number, subjects):
	xfile = openpyxl.load_workbook(filename)
	sheet = xfile.get_sheet_by_name('Judeo-Arabic')
	write_to_cell_number = "AD" + str(current_number)
	#print(sheet[write_to_cell_number].value)
	sheet[write_to_cell_number] = subjects
	xfile.save(FNAME)


example_title = "[al-Tuwarikh al-Yisraʼiliya]."

def assemble_url(title):
	title= title.replace('[', ' ').replace(']', ' ').replace('.', ' ').replace(':', ' ').replace("ʻ", "").replace("(", "").replace(")", "").replace("ʼ", "").replace("'", "").lower().split()
	baseurl = "https://www.worldcat.org/title/"
	title_concat_string = ''

	for word in title:
		title_concat_string += word + "-"
	title_concat_string = title_concat_string[:-1].replace("--", '-')
	primary_url = baseurl + title_concat_string
### secondary URL if the first word in title is "al"
	if title_concat_string[:3] == "al-":
		secondary_url = baseurl + title_concat_string[3:]
		return primary_url, secondary_url
	else:
		return primary_url, None

example_assembled = assemble_url(example_title)
print(example_assembled)

## things to do:
# if starts with al, make a contingency url, that eliminates it DONEZO
# if lands on a page with many items, choose the first one, look for subjects
# if not, choose the second link 

# run the loop again, but if if the value of the subject_tags column is none, then get the info


def make_request(primary_url, secondary_url):
	print("Making a request for new data...")
	resp = requests.get(primary_url)
	resp = resp.content
	soup = BeautifulSoup(resp, 'html.parser')

### if WorldCat brings up error that no results found with primary_url, then try secondary_url
	if soup.find_all(class_ = "error-results", id = "div-results-none"):
		print("found an error the first time")

		### if there is a secondary url to try, try it. 
		if secondary_url != None:
			print(secondary_url)
			resp = requests.get(secondary_url)
			resp = resp.content
			soup = BeautifulSoup(resp, 'html.parser')
			if soup.find_all(class_ = "error-results", id = "div-results-none"):
				print("Found an error with the first and second url!")
				request_without_error = False
		
		## meaning there is no secondary url, we are out of luck with this item 
		else:
			request_without_error = False

	# meaning the primary url worked fine		
	else:
		request_without_error = True
		print('passed')

make_request(*example_assembled)


	# print(resp.status_code)
	# if resp.status_code == 200:
	# 	resp = resp.content
	# 	soup = BeautifulSoup(resp, 'html.parser')
	# 	try:
	# 		subject_div = soup.find_all(id = "subject-terms")
	# 		subject_list = []
	# 		subject_string = ""

	# 		for line in subject_div:
	# 			subjects = line.text.replace('--', "").replace(".", "").split()

	# 			for item in subjects:
	# 				if item not in subject_list:
	# 					subject_list.append(item)
	# 			subject_string = "; ".join(subject_list)
	# 			return subject_string
	# 	except:
	# 		print("### COULDN'T GET SUBJECTS ###")
	# 		return False			
	# else:
	# 	print("### BAD STATUS CODE ###")
	# 	return False






def iterate_excel_file(FNAME):
	wb = load_workbook(filename = FNAME, read_only = True)
	sheet = wb['Judeo-Arabic']
	current_number = 2  ## to start with 
	count_times_written = 0
	count_times_NOT_written = 0 
	title_cell_number = "H" + str(current_number)

	while sheet[title_cell_number].value != None:
		print(title_cell_number)
		current_title = sheet[title_cell_number].value  

		##pass the title into assemble_url, which makes it into a url can make request with
		assembled_url = assemble_url(current_title)

		##making the request with assembled URL, if there's an error, will return false
		result = make_request(assembled_url) 

		if result != False:
			subjects = result
			write_to_excel(FNAME, current_number, subjects)
			count_times_written +=1 
		else:
			count_times_NOT_written += 1

		current_number += 1
		title_cell_number = "H" + str(current_number)

	print("Times written = ")
	print(count_times_written)
	print("Times NOT written = ")
	print(count_times_NOT_written)

	return sheet

#iterate_excel_file(FNAME)





# example_title = "[Aseret ha-devarim] divre Elohim hayim ha-neemarim...ba-ir Tunis... : targum ve-Aravi laTorah"

# example_title = "Shulḥan ʻarukh Oraḥ ḥayim bil-ʻArabi : [ʻim perush Kaf ha-ḥayim ... siman 157[-231]"

# gold_standard = "https://www.worldcat.org/title/aseret-ha-devarim-divre-elohim-hayim-ha-neemarim-ba-ir-tunis-targum-ve-aravi-latorah/"









# url = assemble_url(example_title)
# soup = make_request(url)




def assign_class(loaded_text):
	pass


def get_given_title(FNAME):
	# with open(FNAME, 'rb') as csvfile:
	# 	contents = csv.reader(csvfile)
	# 	next(contents)  # skip header
	# 	data = [r for r in reader]
	# 	print(data[0])


		# for column in csvfile:
		# 	print(', '.join(str(column)))
	pass

def close_excel():
	pass








