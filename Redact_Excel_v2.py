#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Fri Jun 18 11:51:58 2021

@author: gibson
"""

### Program redacts or changes excel fields for given directory of excel files
### this version adds clarity and supports excel files with multiple sheets 
### not recently tested on windows

import os
import pandas as pd
from pandas import Series
import shutil
  


# python script isn't located in directory by default
def config():
    is_Mac = input("Are you running Mac or Windows OS (Mac, Windows):  ")
    if is_Mac == "Windows":
        Path = input("Enter file directory path of files to change (must be the same as this file), for example, 'c:/Users\gibson\Desktop\folder\': ")
        Directory = os.listdir(Path)
        #Directory.remove("Redact_Excel.py")
        #Directory.remove(".git")
        #Directory.remove("REAMDME.md")
        NewDirect = input("Enter the new directory for your files (include \ at end): ") 
        print()
        return Directory, NewDirect, Path
    else:
        Path = input("Enter file directory path of files to change (include / at end): ")
        Directory = os.listdir(Path)
        #Directory.remove("Redact_Excel.py")
        Directory.remove(".DS_Store") ### DS_store may be a hidden file 
        #Directory.remove(".git")
        #Directory.remove("README.md")
        NewDirect = input("Enter the new directory for your files (include / at end): ")
        print()
        return Directory, NewDirect, Path

Directory, New_Direct_Path, Path = config()

Original_Sheets = {}
i = -1
print("Files in program's directory: ")
print()
for items in Directory:
    print(items) # may need to tweek removing hidden items in directory (see what program recognizes)
    i += 1
    # read files
    Original_Sheets[i] = pd.read_excel(Path + items, engine = "openpyxl", sheet_name=None)


def Create_Key_Pairs():
    ''' function takes input and output items for changes, and gives option to take a column input of values to change'''
    Keys=""
    Pairs=""
    PairsFlag = False
    x = input("would you like to change values from a pre-existing column? (y/n): ")
    if x == "y":
         file_p = input("please type file path: ")
         column1 = input("please input column name: ")
         Keys = pd.read_excel(file_p, engine = "openpyxl")
         Keys = Keys[column1]
         Pairs = input("what value would you like to change the fields in this column to? ")
         p_list = []
         for i in range(len(Keys)):
             p_list.append(Pairs)
         Pairs = Series(p_list)
         PairsFlag = True
    elif x == "n":
        k = []
        p = []
        flag = 0
        print()
        print("please enter value to redact followed by ', [corresponding value]'")
        print("when done, enter 'done'")
        while flag == 0:
            kp = input(": ")
            if kp == "done":
                flag = 1
            else:
                key, pair = kp.split(", ")
                k.append(key)
                p.append(pair)
            Keys = Series(k)
            Pairs = Series(p)
    return Keys, Pairs, PairsFlag
        
Keys, Pairs, PairsFlag = Create_Key_Pairs()

mapping_df = pd.concat([Keys, Pairs], axis=1)
mapping_df.columns = ["Name", "NewName"]


def sheet_update(Old_Sheet, mapping_df, PairsFlag):
    ''' function called in create_new_tables() iterations to update one sheet'''
    New_Sheet = '' # avoid reference error (return variable name maintains clarity)
    for i, record in Old_Sheet.iterrows():
        for cell in record:
            if cell in set(Keys):
                if PairsFlag == True:
                    New_Sheet = Old_Sheet.replace(cell, Pairs[0])
                else:
                # gets the index location of keypair table then uses that to replace value 
                    i2 = mapping_df.loc[mapping_df["Name"] == cell].index[0]
                    New_Sheet = Old_Sheet.replace(cell, mapping_df.at[i2, "NewName"])
    return New_Sheet


def create_new_tables(Original_Sheets, mapping_df):
    ''' creates a dictionary of the updated sheets within workbooks'''
    for book in range(0, len(Original_Sheets)):
        for sheet in Original_Sheets[book]:
            Old_Sheet = Original_Sheets[book][sheet]  
            Original_Sheets[book][sheet] = sheet_update(Old_Sheet, mapping_df, PairsFlag)
            New_Sheet_Dict = Original_Sheets
    return New_Sheet_Dict

New_Sheet_Dict = create_new_tables(Original_Sheets, mapping_df)

### must be saved to new directory for multiple sheet version of program
### takes config objects
def save_new_files():
    ''' saves updated xlsx files to new directory'''
    for i in range(0, len(Directory)):
        OldFileNameXLSX = Directory[i]
        OldFileName = OldFileNameXLSX.split(".xlsx")[0]
        NewFileNameXLSX = OldFileName + "_new" + ".xlsx" # optional
        NewFilePath = New_Direct_Path + NewFileNameXLSX
        shutil.copy(Path + OldFileNameXLSX, NewFilePath)
### limitation: default ExcelWriter engine doesn't auto format as a table in excel
        with pd.ExcelWriter(NewFilePath) as writer:    
            for sheetname in New_Sheet_Dict[i]:
                df = New_Sheet_Dict[i][sheetname]
                df.to_excel(writer, sheet_name=sheetname, index=False)
    return

save_new_files()


