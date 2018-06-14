import json
import requests
import csv
import secrets as secrets


api_key = secrets.api_key


##later can add more self definitions for other categories
class Document:
	def __init__(self, title):
		self.title = title
		self.subject_tags = ""

	def __str__(self):
		return subject_tags
	pass


def get_params():
	pass


def make_request(baseurl, params):
	print("Making a reuqest for new data...")
	resp = requests.get(baseurl, params)
	loaded_text = json.loads(resp.text)
	return loaded_text
	pass


def assign_class(loaded_text):
	pass


def get_given_title():

	return title


def open_excel():
	with 

def close_excel():
	pass


def get_given_title():
	csv = open_excel()
	for document_item in csv:





