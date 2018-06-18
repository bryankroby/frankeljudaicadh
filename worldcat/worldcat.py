import json
import requests
import csv
import secrets as secrets
from bs4 import BeautifulSoup
import pandas as pd
import re

from openpyxl import *

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

def assemble_url(title):
	title= title.replace('[', ' ').replace(']', ' ').replace('.', ' ').replace(':', ' ').replace("ʻ", "").replace("(", "").replace(")", "").split()
	baseurl = "https://www.worldcat.org/title/"
	title_concat_string = ''
	for word in title:
		title_concat_string += word + "-"

	title_concat_string = title_concat_string[:-1].replace("--", '-')
	complete_url = baseurl + title_concat_string
	print(complete_url)
	return complete_url

def make_request(complete_url):
	print("Making a request for new data...")
	resp = requests.get(complete_url)

	print(resp.status_code)
	if resp.status_code == 200:
		resp = resp.content
		soup = BeautifulSoup(resp, 'html.parser')
		subject_div = soup.find_all(id = "subject-terms")
		subject_list = []

		for line in subject_div:
			subjects = line.text.replace('--', "").replace(".", "").split()

			for item in subjects:
				if item not in subject_list:
					subject_list.append(item)
			return subject_list
	else:
		print("### BAD STATUS CODE ###")
		return False


def iterate_excel_file(FNAME):
	wb = load_workbook(filename = FNAME)
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

iterate_excel_file(FNAME)





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








