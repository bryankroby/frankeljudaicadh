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



# def load_excel_file(FNAME):
# 	xl = pd.ExcelFile(FNAME)
# 	df = xl.parse('Judeo-Arabic')
# 	return df

def load_excel_file(FNAME):
	wb = load_workbook(filename = FNAME)
	sheet_ranges = wb['Judeo-Arabic']



	print(sheet_ranges['H18'].value)


#load_excel_file(FNAME)	

#current_number = which cell we're on

current_number = 2  ## to start with 
cell_number = "H" + str(current_number)
# print(cell_number)
current_number += 1



def write_to_excel(filename, cell_number):
	xfile = openpyxl.load_workbook(filename)
	sheet = xfile.get_sheet_by_name('Judeo-Arabic')

	print(sheet[cell_number].value)
	sheet["AC2"] = "THIS is a test!"
	xfile.save(FNAME)

write_to_excel(FNAME, cell_number)



	# xl = load_workbook(FNAME)


	# df = xl.parse('Judeo-Arabic')
	# return df




# dataframe = load_excel_file(FNAME)
# for column in dataframe["Title"]:
# 	print(column)






# example_title = "[Aseret ha-devarim] divre Elohim hayim ha-neemarim...ba-ir Tunis... : targum ve-Aravi laTorah"

# example_title = "Shulḥan ʻarukh Oraḥ ḥayim bil-ʻArabi : [ʻim perush Kaf ha-ḥayim ... siman 157[-231]"

# gold_standard = "https://www.worldcat.org/title/aseret-ha-devarim-divre-elohim-hayim-ha-neemarim-ba-ir-tunis-targum-ve-aravi-latorah/"



def assemble_url(title):
	title= title.replace('[', ' ').replace(']', ' ').replace('.', ' ').replace(':', ' ').replace("ʻ", "").split()
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








