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

#opens excel file, writes subjects to the given subject_tag cell number --> WILL NEED TO CHANGE SPECIFIED CELL # LATER
def write_to_excel(filename, current_number, subjects):
	xfile = openpyxl.load_workbook(filename)
	sheet = xfile.get_sheet_by_name('Judeo-Arabic')
	write_to_cell_number = "AE" + str(current_number)
	#print(sheet[write_to_cell_number].value)
	sheet[write_to_cell_number] = subjects
	xfile.save(FNAME)


#takes give title, eliminates uncessary characters and tokenizes, creates primary url. 
#if first word in title is "al" creates, secondary url
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


##try to find subject div, if doesn't exist, meaning there are no subjects, return false
##if there are subjects, return subject string
def find_subjects(soup):
	#if the item has subject terms
	if soup.find_all(id = "subject-terms"):
		subject_div = soup.find_all(id = "subject-terms") 
		subject_list = []
		subject_string = ""

		for line in subject_div:
			subjects = line.text.replace('--', "").replace(".", "").split()

			for item in subjects:
				if item not in subject_list:
					subject_list.append(item)
			subject_string = "; ".join(subject_list)
			print(subject_string)
			return subject_string
	#the item doesn't have subject terms 
	else:
		##couldn't get subjects 
		return False			


## if not error, need to check if on a menu page or if on the item's page
## if on the menu page, follow the link into the first menu, get subject_tags, as usual
## possibly investigate second href on the menu page, as well

## improve what happens with tags with commas in them like:
#"Esther,; Queen; of; Persia; Bible; stories,; Judeo; Arabic; Esther; Old; Testament"
# or " Judeo-Arabic; literature; Jews; History; To; 70; AD"


def make_soup(url):
	print("Making a request for new data...")
	resp = requests.get(url)
	resp = resp.content
	soup_contents = BeautifulSoup(resp, 'html.parser')
	return soup_contents


##tries to make request with two different URLS and looks for subjects
def make_request(primary_url, secondary_url):
	soup = make_soup(primary_url)

	## if WorldCat brings up error that no results found with primary_url, then try secondary_url
	if soup.find_all(class_ = "error-results", id = "div-results-none"):
		print("found an error with the first url")

		### if there is a secondary url to try, try it. 
		try: 
			soup = make_soup(secondary_url)
			if soup.find_all(class_ = "error-results", id = "div-results-none"):
				print("Found an error with the first and second url!")
				##  secondary_url returned error, as well
				request_without_error = False

			## secondary_url worked!	
			else:
				request_without_error = True

		### if there is no secondary url, out of luck, ultimately false
		except:
			request_without_error = False

	# meaning the primary url worked fine		
	else:
		request_without_error = True

	## no error with URLs, now find out if on menu page and look for subjects 
	if request_without_error == True:

		##need to find out if on menu page or specific item page
		the_menu_exists = soup.find(class_ = "menuElem")
		print("the the_menu_exists")
		## if we are indeed on the menu page... 
		if the_menu_exists:
			baseurl = "https://www.worldcat.org"

			menu_items = the_menu_exists.find_all(class_= "result details")
			print("the menu items exist")

			href = menu_items[0].find("a")['href']

			complete_url = baseurl + href


			print(complete_url)
			print("#################")
			##need to follow the first href and get its soup


		## will return the subject string if exists, if not, returns False
		subject_string_or_false = find_subjects(soup)
		print(subject_string_or_false)
		return subject_string_or_false
	#returns false if not able to reach an appropriate URL
	else:
		return request_without_error


example_assembled = 'https://www.worldcat.org/title/kohelet-im-sharh-ha-arvi-ha-meduberet-ben-ha-am-ve-im-perush-shema-shelomoh'
# print(example_assembled)
made = make_request(example_assembled, None)
print(made)



def iterate_excel_file(FNAME):
	wb = load_workbook(filename = FNAME, read_only = True)
	sheet = wb['Judeo-Arabic']
	current_number = 2  ## to start with 
	count_times_written = 0
	count_times_NOT_written = 0 
	title_cell_number = "H" + str(current_number)
	print("Title cell number:")
	print(title_cell_number)
	write_to_cell_number = "AE" + str(current_number)

	#for all the rows of items
	while sheet[title_cell_number].value != None:

		#if item does not already have subjects assined to it
	#		print("value of subject cell is None!!!")
		#print(title_cell_number)
		current_title = sheet[title_cell_number].value  

		##pass the title into assemble_url, which makes it into a url can make request with
		assembled_url = assemble_url(current_title)

		##making the request with assembled URL, if there's an error, will return false
		result = make_request(*assembled_url) 

		## meaning it was able to find subjects
		if result != False:
			subjects = result
			write_to_excel(FNAME, current_number, subjects)
			count_times_written +=1 
		else:
			count_times_NOT_written += 1

		current_number += 1
		title_cell_number = "H" + str(current_number)

		#there is already a subject assigned to the cell number 
		# else:
		# 	continue

	print("Times written = ")
	print(count_times_written)
	print("Times NOT written = ")
	print(count_times_NOT_written)

	#return sheet

#iterate_excel_file(FNAME)



## things to do:
# if starts with al, make a contingency url, that eliminates it DONEZO
# if not, choose the second link  DONEZO
# run the loop again, but if if the value of the subject_tags column is none, then get the info DONEZO

# if lands on a page with many items, choose the first one, look for subjects

