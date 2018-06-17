import json
import requests
import csv
import secrets as secrets
from bs4 import BeautifulSoup
import pandas as pd
import re

# api_key = secrets.api_key



# FNAME = "master_copy_harvard_hebraic_collection_61418.csv"


FNAME = "Judeo_arabic_spreadsheet.xlsx"

def load_excel_file(FNAME):
	xl = pd.ExcelFile(FNAME)
	df1 = xl.parse('Judeo-Arabic')
	return df1
	# for a_title in df1['Title']:
	# 	print(a_title)


##later can add more self definitions for other categories
class Document:
	def __init__(self, title):
		self.title = title
		self.subject_tags = ""

	def __str__(self):
		return subject_tags
	pass


# example 
#https://www.worldcat.org/title/aseret-ha-devarim-divre-elohim-hayim-ha-neemarim-ba-ir-tunis-targum-ve-aravi-latorah


example_title = "[Aseret ha-devarim] divre Elohim hayim ha-neemarim...ba-ir Tunis... : targum ve-Aravi laTorah"


gold_standard = "https://www.worldcat.org/title/aseret-ha-devarim-divre-elohim-hayim-ha-neemarim-ba-ir-tunis-targum-ve-aravi-latorah/"
example_title = example_title.replace('[', ' ').replace(']', ' ').replace('.', ' ').replace(':', ' ').split()



def assemble_url(title):
	baseurl = "https://www.worldcat.org/title/"
	title_concat_string = ""
	for word in title:
		title_concat_string += word + "-"
	title_concat_string = title_concat_string[:-1]
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
			subjects = line.text.replace('-', "").replace(".", "").split()

			for item in subjects:
				if item not in subject_list:
					subject_list.append(item)
			print(subject_list)

	else:
		print("### BAD STATUS CODE ###")



url = assemble_url(example_title)

soup = make_request(url)






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








