### Script to find all the .xls or .xlsx EDD's in a folder and stack them for upload as one EDD

# takes all the .xlsx or .xls files in a folder and stacks them into one edds

# - stack them and export as new file


#####  HOW TO USE ME
# DROP ME IN A FOLDER WITH THE EXCEL FILES YOU WISH TO STACK
# OPEN UP COMMAND PROMPT AND CD TO THE folder

# TYPE python stack_edds.py and hit enter
##############################################################################################

import os
import pandas as pd
import numpy as np


# find all the files in the folder
def find_excel():
    # finds all the .xls or .xlsx files in the folder
    files = []
    for file in os.listdir():
        if file.endswith(".xlsx") or file.endswith(".xls"):
            files.append(file)

    # returns a list of the file names
    return files

# load each file in the list into a list as a pandas df
def excel_to_df(files):
    # loads each file in the input list into a list of dictionaries
    dfs = []
    for file in files:
        dfs.append(pd.read_excel(file,sheet_name=None))

    return dfs

def check_tab_names(dfs):
    # checks to make sure the tab names are all the same for each dictionary
    keys = []
    for df in dfs:
        keys.append(df.keys())
    # check that the strings of keys are the same
    keysgood = keys[1:] == keys[:-1]

    if keysgood is False:
        try:
            t = 0/0
        except:
           raise("One or More Sets of Excel Tabs Have Different Names Rename Them Before Proceeding ")

    return keysgood,keys

def combine_tabs(dfs,keys):

    tabs=[]
    for key in keys:
        for k in key:
            tabs.append(k)
    unqkeys = np.unique(tabs)

    # loop through each file in the folder, structure the similar tabs as a pandas df in a dictionary of tab unique names
    edds = {}
    for key in unqkeys:
        edds[str(key)] = []
        for file in os.listdir():
            if file.endswith(".xlsx") or file.endswith(".xls"):
                df = pd.read_excel(file,sheet_name=key)
                # get rid of completley null rows bc sometimes it loads at 65000 black excel rows below data
                df.dropna(how='all',inplace=True)
                edds[key].append(df)
      

    # Create a Pandas Excel writer using XlsxWriter as the engine.
    writer = pd.ExcelWriter('Combined_EDDs.xlsx', engine='xlsxwriter')

    # finally combine into seperate tabs
    final =[]
    for tab in edds:
        # get rid of completley null rows bc sometimes it loads at 65000 black excel rows below data

        r = pd.concat(edds[tab])


        final.append(r)
        #print (r.shape())
        r.to_excel(writer,sheet_name=tab,index=False)
    writer.save()

if  __name__  == '__main__':
    files = find_excel()
    print("Combining the following files:")
    for f in files:
        print(f)
    dfs = excel_to_df(files)
    good,keys = check_tab_names(dfs)
    print("Combining tabs...")
    combine_tabs(dfs,keys)
    print('File saved to Combined_EDDs.xlsx')



##### Notes: to many loops and appends but I dont think this will really matter time wise
# compared to the actual entire process of loading anyway.
