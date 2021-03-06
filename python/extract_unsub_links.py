# Extract Unsubscribe Links
# By Vanesa Hercules, 2022

import gspread
from bs4 import BeautifulSoup
import re
import pandas as pd

# Assign Service Account Key
gc = gspread.service_account(filename='client_secret.json')

# Open Google Sheet
sh = gc.open("Unsubscribe Gmail Messages")

# Get All Values From Body Column
values_list = sh.sheet1.col_values(4)

# Create List of Output Cells Per Length of value_ list
cell_list = []
for cell in range(1, len(values_list)+1):
    cell_str = str(cell)
    cell_list.append('E'+ cell_str) 

# Zip cell_list with values_list into a DataFrame
df = pd.DataFrame(list(zip(cell_list, values_list)), columns =['cell', 'body']) 

# Search for unsubscribe links and write them back to Google Sheet
def search_for_unsublink (body):
    """ Searches for 'unsubscribe' text among all a tags and returns href if found or none"""
    soup = BeautifulSoup(body, features='lxml')
    a_list = list(soup.find_all('a'))
    for a in a_list:
        if re.search('nsubscribe', str(a)):
            unsubscribe_a = a
            unsubscribe_link = unsubscribe_a.get('href')
            return unsubscribe_link

def write_link_to_cell(cell, link):
    """ Write links to Google Sheet"""
    sh.sheet1.update(cell, link)

# Print iterations progress
def printProgressBar (iteration, total, prefix = '', suffix = '', decimals = 1, length = 100, fill = '█', printEnd = "\r"):
    """
    Call in a loop to create terminal progress bar
    @params:
        iteration   - Required  : current iteration (Int)
        total       - Required  : total iterations (Int)
        prefix      - Optional  : prefix string (Str)
        suffix      - Optional  : suffix string (Str)
        decimals    - Optional  : positive number of decimals in percent complete (Int)
        length      - Optional  : character length of bar (Int)
        fill        - Optional  : bar fill character (Str)
        printEnd    - Optional  : end character (e.g. "\r", "\r\n") (Str)
    """
    percent = ("{0:." + str(decimals) + "f}").format(100 * (iteration / float(total)))
    filledLength = int(length * iteration // total)
    bar = fill * filledLength + '-' * (length - filledLength)
    print(f'\r{prefix} |{bar}| {percent}% {suffix}', end = printEnd)
    # Print New Line on Complete
    if iteration == total: 
        print()

# Main
l = len(df)
for index, row in df.iterrows():
    result = search_for_unsublink(row['body'])
    write_link_to_cell(row['cell'], result)
    # Update Progress Bar
    printProgressBar(index + 1, l, prefix = 'Progress:', suffix = 'Complete', length = 50)
