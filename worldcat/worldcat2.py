import json
import requests
import csv
import secrets as secrets
from bs4 import BeautifulSoup
import pandas as pd
import re

from openpyxl import Workbook 

from openpyxl import load_workbook

import openpyxl

# api_key = secrets.api_key

FNAME = "Judeo_arabic_spreadsheet copy 2.xlsm"


## creating a class for each publication
class Publication(object):
    def __init__(self, title):
        self.title = title
        self.subject_tags = ""
    def __str__(self):
        return subject_tags
    pass

#opens excel file, writes subjects to the given subject_tag cell number --> WILL NEED TO CHANGE SPECIFIED CELL # LATER
def write_to_excel(write_to_cell_number, content_to_write):
    wb = Workbook()
    xfile = openpyxl.load_workbook(filename = FNAME, read_only=False, keep_vba=True)
    sheet = xfile.get_sheet_by_name('Judeo-Arabic')
    #print(sheet[write_to_cell_number].value)
    sheet[write_to_cell_number].value = content_to_write
    xfile.save(FNAME)


## perform reg ex on a term that is going to be used in the URL query
def reg_ex(query_term):
    improved_query_term = query_term.replace('[', ' ').replace(']', ' ').replace('.', ' ').replace(':', ' ').replace("ʻ", "").replace("(", "").replace(")", "").replace("ʼ", "").replace("'", "").replace(",", "").lower().split()
    complete_query_string = ''
    for word in improved_query_term:
        complete_query_string += word + "-"
    complete_query_string = complete_query_string[:-1].replace("--", '-')

    return complete_query_string

#creates request object and parses w/ BeautifulSoup
def make_soup(url):
    resp = requests.get(url)
    resp = resp.content
    soup_contents = BeautifulSoup(resp, 'html.parser')
    return soup_contents


#takes give title, eliminates uncessary characters and tokenizes, creates primary url. 
#if first word in title is "al" creates, secondary url
def assemble_url(a_title, an_author):
    title= reg_ex(a_title)
    baseurl = "https://www.worldcat.org/search?q=" + title
    primary_url = baseurl
    ## if there is an author:  need to get rid of the title part in the base url

### getting rid of this cause causes problems if not the same spelling of author..... might run it both ways and see the error
    # author = reg_ex(an_author)
    # primary_url = baseurl + "-" + author
    return primary_url


##try to find subject div, if doesn't exist, meaning there are no subjects, return false and record the error in error_cell_number
##if there are subjects, return subject string. 
def find_subjects(soup, error_cell_number):
    #if the item has subject terms
    if soup.find_all(id = "subject-terms"):
        subject_div = soup.find_all(id = "subject-terms")
        # subject_div = subject_div.find_all('li', attrs={'class' :"subject-term"})
        subject_list = []
        subject_string = ""

        for line in subject_div:
            subjects = line.text.replace('--', "").replace(".", "").split()

            for item in subjects:
                if item not in subject_list:
                    subject_list.append(item)
            subject_string = "; ".join(subject_list)
            #print(subject_string)

            subject_string.replace("; View; all; subjects", "")
            return subject_string
    #the item doesn't have subject terms 
    else:
        ##couldn't get subjects 
        print("Could not find subjects")
        error_message = "No subjects"
        write_to_excel(error_cell_number, error_message)
        return False            


## finds and returns description or false. does not record an error message
def find_page_number_description(soup):
    try:
        description_div = soup.find(id = "details-description")
        description = description_div.find("td").text

        print(description)
        return description
    except:
        print("no description")
        return False



## finds and returns notes or false. does not record an error message
def find_item_notes(soup):
    # if notes_div exists:
    try:
        notes_div = soup.find(id = "details-notes")
        notes = notes_div.find("td")
        notes = str(notes).replace("<br/>", " ").replace("<td>", "").replace("</td>", "")
        print(notes)
        return notes
    # no notes_div exists
    except:
        print("no notes")
        return False


## makes search request  looks for subjects
def make_request(primary_url, error_cell_number):
    print("Making request for new data using primary_url")
    soup = make_soup(primary_url)

    ## if WorldCat brings up error that no results found with primary_url, then try secondary_url
    if soup.find_all(class_ = "error-results", id = "div-results-none"):
        error_message = "Error with primary URL"
        write_to_excel(error_cell_number, error_message)
        print(error_message)
        request_without_error = False

    ## no error with primary_URL, now find out if on menu page and look for subjects 
    else:
        request_without_error = True
        ##need to find out if on menu page or specific item page
        baseurl = "https://www.worldcat.org"
        
        ## follow the first href on menu page and get its soup
        menu = soup.find(class_ = "menuElem")
        menu_items = menu.find_all(class_= "result details")
        print("The menu items exist")


        ## verify in the correct language in the first link
        verifying_language = menu_items[0].find(class_= "itemLanguage").text
        if verifying_language == "Judeo-Arabic" or "Hebrew":
            href = menu_items[0].find("a")['href']
            complete_url = baseurl + href
            print("Making request for new data using HREF from 1st in menuElem")
            item_page_soup = make_soup(complete_url)
            valid_entries = True
        ## check out the language in the second link, if no second link, the except 
        elif menu_items[1].find(class_= "itemLanguage").text == "Judeo-Arabic" or "Hebrew":
            try:            
                verifying_language = menu_items[1].find(class_= "itemLanguage").text
                if verifying_language == "Judeo-Arabic" or "Hebrew":
                    href = menu_items[1].find("a")['href']
                    complete_url = baseurl + href
                    print("Making request for new data using HREF from 2nd in menuElem")
                    item_page_soup = make_soup(complete_url)
                    valid_entries = True
            # meaning, there is no second link on menu
            except:
                valid_entries = False

        ## neither of the first two entries in judeo arabic or hebrew, or there is only one entry and it isn't in JA or H:
        else:
            valid_entries = False
            error_message = "Wrong language"
            write_to_excel(error_cell_number, error_message)
            print(error_message)            
            return valid_entries   ### --> in this case, only thing returned is false.


    # else:
    #     pass

        ## able to find a valid entry, will look for and return the subject string if exists, if not, returns False
        if valid_entries == True:
            subject_string_or_false = find_subjects(item_page_soup, error_cell_number)
            item_description_or_false = find_page_number_description(item_page_soup)
            item_notes_or_false = find_item_notes(item_page_soup)


        ### need to find number of pages, and other notes, if possible, and return a tuple 
        return_tuple = (subject_string_or_false, item_description_or_false, item_notes_or_false)
        print(len(return_tuple))
        print(return_tuple)
        return return_tuple



def iterate_excel_file():
    wb = load_workbook(filename = FNAME, read_only = True)
    ## getting a specific sheet from the XL
    sheet = wb['Judeo-Arabic']
    sheet = wb.active   ## got this from https://stackoverflow.com/questions/49159245/python-error-on-get-sheet-by-name
    
    ## Cell row to start with 
    # current_number = 2  ############ ---> Original start row = 2


    current_number = 13

    title_cell_number = "H" + str(current_number)
    author_cell_number = "G" + str(current_number)
    subject_cell_number = "W" + str(current_number)
    description_cell_number = "X" + str(current_number)
    notes_cell_number = "Y" + str(current_number)
    error_cell_number= "Z" + str(current_number)

    #for all the rows of items
    while sheet[title_cell_number].value != None:


    # for i in range(3):
        #if item does not already have subjects assined to it
    #       print("value of subject cell is None!!!")
        print("Title cell number:")
        print(title_cell_number)
        current_title = sheet[title_cell_number].value


        ### making sure the title is more than one word long
        if len(current_title) > 1:
            print(current_title)
            current_author = sheet[author_cell_number].value
            print(current_author)
            ##pass the title and author into assemble_url, which makes it into a url can make request with
            if current_author != "":
                assembled_url = assemble_url(current_title, current_author)
                print(assembled_url)
            else:
                assembled_url = assembled_url(current_title)

    ## could also make it so that it always puts in the the author, if putting in a "" author doesn't mess it up. 
            ##making the request with assembled URL, if there's an error, will return false
            result = make_request(assembled_url, error_cell_number) 

            ## meaning there was a valid entry 
            # if result != False:

            # <1 would mean  was no valid entry and therefore, didn't get subjects
            ## meaning there was a valid entry and it went forth to check subjects, notes, description
            if result != False:
            # if len(result) >1:
                subjects = result[0]
                description = result[1]
                notes = result[2]

                #it was able to find subjects
                if subjects != False:
                    write_to_excel(subject_cell_number, subjects)

                #if found a description
                if description != False:
                    write_to_excel(description_cell_number, description)

                #if found notes
                if notes != False:
                    write_to_excel(notes_cell_number, notes)

        ## writing in excel saying it was skipped because of non-distinct name
        else:
            error_message = "Indistinct title"
            write_to_excel(error_cell_number, error_message)


        print("\n###################\n")
        current_number += 1

        ## increase the cell numbers by 1
        title_cell_number = "H" + str(current_number)
        author_cell_number = "G" + str(current_number)
        subject_cell_number = "W" + str(current_number)
        description_cell_number = "X" + str(current_number)
        notes_cell_number = "Y" + str(current_number)
        error_cell_number= "Z" + str(current_number)



    #return sheet




######### To demo a part of the script #######
# example_assembled = 'https://www.worldcat.org/title/kohelet-im-sharh-ha-arvi-ha-meduberet-ben-ha-am-ve-im-perush-shema-shelomoh'

# example_assembled = "https://www.worldcat.org/title/-20160603/oclc/6914317075&referer=brief_results"
# made = make_request(example_assembled, None)
# print(made)


######### TO RUN THE SCRIPT, UNCOMMENT BELOW! ######
iterate_excel_file()



######  SCRIPT TO-DO LIST #####
# if starts with al, make a contingency url, that eliminates it DONEZO
# if not, choose the second link  DONEZO
# run the loop again, but if if the value of the subject_tags column is none, then get the info DONEZO
# if lands on a page with many items, choose the first one, look for subjects DONEZO

## possibly investigate second href on the menu page, as well

## improve what happens with tags with commas in them like:
#"Esther,; Queen; of; Persia; Bible; stories,; Judeo; Arabic; Esther; Old; Testament"
# or " Judeo-Arabic; literature; Jews; History; To; 70; AD"


### make a contingency before getting something from the menu that it is in hebrew or judeo-arabic

## need to make it so file format doesn't die after writing on it.


#instead of constructing url of title, make a url of title and author, if available 
#also, confirm the text is in hebrew / j-a


### add a flagging to the spread sheet in a different column to say there was an error, couldn't find


