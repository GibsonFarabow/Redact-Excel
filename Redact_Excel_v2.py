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
import shutil
  


def orig_excel_sheets():
    """returns a list of string directory files and the directory path, and imports excel files through a pandas read"""
    def os_config():
        # python script isn't located in directory by default
        # os_config may need to be updated to user needs
        op_sys = input("Are you running Mac or Windows OS (Mac, Windows):  ")
        if op_sys == "Windows":
            Direct_Path = input("Enter file directory path of files to change (must be the same as this file), for example, 'c:/Users\gibson\Desktop\folder\': ")
            Directory = os.listdir(Path)
            #Directory.remove("Redact_Excel.py")
            #Directory.remove(".git")
            #Directory.remove("REAMDME.md")
            print()
            return Directory, Direct_Path
        else:
            Path = input("Enter file directory path of files to change (include / at end): ")
            Directory = os.listdir(Path)
            #Directory.remove("Redact_Excel.py")
            Directory.remove(".DS_Store") # DS_store may be a hidden file 
            #Directory.remove(".git")
            #Directory.remove("README.md")
            print()
            return Directory, Direct_Path
    ###
    Directory,  Direct_Path = os_config()
    Original_Sheets = {}
    print("Files in program's directory: ")
    print()
    i = -1
    for items in Directory:
        print(items) # may need to tweek removing hidden items in directory (see what program recognizes)
        i += 1
        Original_Sheets[i] = pd.read_excel(Path + items, engine = "openpyxl", sheet_name=None)
    return Original_Sheets, Directory, Direct_Path


def Create_Key_Pairs():
    ''' function takes input excel field strings for changes, and gives option to take a column input of values to change'''
    ''' output assists in excel sheet manipulation '''
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
         Pairs = pd.Series(p_list)
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
            Keys = pd.Series(k)
            Pairs = pd.Series(p)
    return Keys, Pairs, PairsFlag


    def sheet_update(Old_Sheet, mapping_df, PairsFlag):
        ''' function called in create_new_tables() iterations to update one sheet'''
        New_Sheet = ''
        for i, record in Old_Sheet.iterrows():
            for cell in record:
                if cell in set(Keys):
                    if PairsFlag == True:
                        New_Sheet = Old_Sheet.replace(cell, Pairs[0])
                    else:
                    # gets the index location of key pair table then uses that to replace value 
                        i2 = mapping_df.loc[mapping_df["Name"] == cell].index[0]
                        New_Sheet = Old_Sheet.replace(cell, mapping_df.at[i2, "NewName"])
        return New_Sheet


def create_new_tables(Original_Sheets, Keys, Pairs, PairsFlag):
    ''' creates a dictionary of the updated sheets within workbooks, filename: [dataframe sheet] '''
        def sheet_update(Old_Sheet, mapping_df, PairsFlag):
            ''' function iterates through and updates one sheet'''
            New_Sheet = ''
            ### revisit for efficiency
            for i, record in Old_Sheet.iterrows(): # pandas function
                for cell in record:
                    if cell in set(Keys):
                        if PairsFlag == True:
                            New_Sheet = Old_Sheet.replace(cell, Pairs[0])
                        else:
                        # gets the index location of key pair table then uses that to replace value 
                            i2 = mapping_df.loc[mapping_df["Name"] == cell].index[0]
                            New_Sheet = Old_Sheet.replace(cell, mapping_df.at[i2, "NewName"])
            return New_Sheet
    ###
    mapping_df = pd.concat([Keys, Pairs], axis=1)
    mapping_df.columns = ["Name", "NewName"]
    for book in range(len(Original_Sheets)):
        for sheet in Original_Sheets[book]:
            Old_Sheet = Original_Sheets[book][sheet]  
            # updates original sheet then saves it as new sheet
            Original_Sheets[book][sheet] = sheet_update(Old_Sheet, mapping_df, PairsFlag)
            New_Sheet_Dict = Original_Sheets
    return New_Sheet_Dict


def save_new_files(Direct_Path, Directory, New_Sheet_Dict):
    ''' saves updated xlsx files to new directory'''
    New_Direct_Path = input('input new file path (include /): ')
    for i in range(len(Directory)):
        OldFileNameXLSX = Directory[i]
        OldFileName = OldFileNameXLSX.split(".xlsx")[0]
        NewFileNameXLSX = OldFileName + "_new" + ".xlsx" # optional
        NewFilePath = New_Direct_Path + NewFileNameXLSX
        shutil.copy(Path + OldFileNameXLSX, NewFilePath)
# limitation: default ExcelWriter engine doesn't auto format as a table in excel unless there is only one sheet
        with pd.ExcelWriter(New_Direct_Path) as writer:    
            for sheetname in New_Sheet_Dict[i]:
                df = New_Sheet_Dict[i][sheetname]
                df.to_excel(writer, sheet_name=sheetname, index=False)
    return


### imports files through read and obtains os path and string item names of directory items, ex: "README.md"
Original_Sheets, Directory, Direct_Path = orig_excel_sheets()

### assist in dataframe manipualtion
Keys, Pairs, PairsFlag = Create_Key_Pairs()

### creates dictionary of file name path keys with pairs of dataframes of each sheet
New_Sheet_Dict = create_new_tables(Original_Sheets, Keys, Pairs, PairsFlag)

### saves updated files to your system, either as _new files or replacing the old
save_new_files(Direct_Path, Director, New_Sheet_Dict)


