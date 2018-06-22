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



resp = requests.get("https://www.worldcat.org/title/sefer-hilkhot-ha-rif-al-masekhet-pesah-rishon-pesahim-perakim-1-4-10/oclc/23531150&referer=brief_results#relatedsubjects")
resp = resp.content
soup = BeautifulSoup(resp, 'html.parser')

#if the item has subject terms
subject_div = soup.find_all(id = "subject-terms-detailed")
subject_list = []
subject_string = ""

for line in subject_div:
    subjects = line.text.replace(".", "").split("\n")
    for a_line in subjects:
        a_line = a_line.split(" -- ")


## finally i've got a line that is all clean
        for word in a_line:
            ## if word isn't empty ''
            if len(word) >0:
            ## getting rid of unnecessary comma cause i'm going to join them with semi-colon
                if word[-1] == ",":
                    new_word = word[:-1] 
                else:
                    new_word = word
                #now, i have the official "new word" either way
                # print(subject_list)
                if new_word not in subject_list:
                    print(new_word + str(len(new_word)))
                    subject_list.append(new_word)

                    y = new_word.encode('utf-8')
                    print(y)

                else:
                    print(new_word + " was already in subject_list")
            # else:
            #     print("len of word wasn't >0")


print(subject_list)


                # print(subject_list)

#                 if word not in subject_list:
#                     print(word)
#                     subject_list.append(word)

# print(subject_list)







                # 

                #     print(word)
                #     
                #     print(subject_list)
            # subject_string = "; ".join(subject_list)
# print(subject_string)
# print(subject_list)


    #     if item not in subject_list:
    #         


# #if the item has subject terms
# if soup.find_all(id = "subject-terms"):
#     subject_div = soup.find_all(id = "subject-terms")
#     # subject_div = subject_div.find_all('li', attrs={'class' :"subject-term"})
#     subject_list = []
#     subject_string = ""

#     for line in subject_div:
#         subjects = line.text.replace('--', "").replace(".", "").split()

#         for item in subjects:
#             if item not in subject_list:
#                 subject_list.append(item)
#         subject_string = "; ".join(subject_list)
        #print(subject_string)


# #if the item has subject terms
# if soup.find_all(id = "subject-terms"):
#     subject_div = soup.find_all(id = "subject-terms")
#     # subject_div = subject_div.find_all('li', attrs={'class' :"subject-term"})
#     subject_list = []
#     subject_string = ""

#     for line in subject_div:
#         subjects = line.text.replace('--', "").replace(".", "").split()

#         for item in subjects:
#             if item not in subject_list:
#                 subject_list.append(item)
#         subject_string = "; ".join(subject_list)
#         #print(subject_string)

#         subject_string.replace("; View; all; subjects", "")




