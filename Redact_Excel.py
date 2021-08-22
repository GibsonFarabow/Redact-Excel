#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Fri Jun 18 11:51:58 2021

@author: gibson
"""

#link excel files
import os
from os import listdir

import pandas as pd
from pandas import DataFrame, Series
  

is_Mac = input("Are you running Mac or Windows OS (Mac, Windows):  ")

def Direct():
    if is_Mac == "Windows":
        Path = input("Enter file directory path of files to change (must be the same as this file), for example, 'c:/Users\gibson\Desktop\folder\': ")
        Directory = os.listdir(Path)
        Directory.remove("Redact_Excel.py")
        NewDirect = input("Enter the new directory for your files (include \ at end), or leave blank input to keep the same: ") 
        return Directory, NewDirect, Path
    else:
        Path = input("Enter file directory path of files to change (must be the same as this file): ")
        Directory = os.listdir(Path)
        Directory.remove(".DS_Store")
        Directory.remove("Redact_Excel.py")
        NewDirect = input("Enter the new directory for your files (include / at end), or leave blank input to keep the same: ") ### check include vs don't include
        return Directory, NewDirect, Path

Directory, NewDirect_Path, Path = Direct()

#Directory = os.listdir('/Users/gibson/Desktop/projects/redact/')
#Directory.remove(".DS_Store")
#Directory.remove("RedactExcel-copy.py")
#Directory.remove("pair.xlsx")


'''
def Is_in_direct_test(Directory):
    if "RedactExcelNew.py" in Directory:
        Directory.remove("RedactExcel.py")
        if is_Mac = "Mac":
            Directory.remove(".DS_Store")
    else:
        print("invalid input, please try again \n")
        Directory = Direct()
    return Directory
'''


Original_Sheets = {}
i = -1
for items in Directory:
    print(items)
    i += 1
    Original_Sheets[i] = pd.read_excel(Path + items, engine = "openpyxl")


def Create_Key_Pairs():
    redact = {}
    Sheets =[]
    keypair = []
    Keys=""
    Pairs=""
    PairsFlag = False
    ### below file would be unchanged in its transformed output
    x = input("would you like to create keys from a pre-existing column? (y/n): ")
    if x == "y":
         file_p = input("please type file path: ")
         column1 = input("please input column name: ")
         Keys = pd.read_excel(file_p, engine = "openpyxl")
         Keys = Keys[column1]
         Pairs = input("if you would like to replace with one value, input y else input n: ")
         if Pairs == "y":
             Pairs = input("is your value a number? ")
             if Pairs == "y":
                 Pairs = int(input("enter number: "))
             else: 
                 Pairs = input("enter value: ")
             p_list = []
             for i in range(len(Keys)):
                 p_list.append(Pairs)
             Pairs = Series(p_list)
             PairsFlag = True
         else:
            y1 = input("please provide corresponding replacement column name: ")
            Pairs = pd.read_excel(file_p, engine = "openpyxl")
            Pairs = Pairs[y1]
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



df = pd.concat([Keys, Pairs], axis=1)
df.columns = ["Name", "NewName"]


def transform(Original_Sheets, df):
    for sheet_key in Original_Sheets:
        A = Original_Sheets[sheet_key]      
        Original_Sheets[sheet_key] = sheets_recursive(A, df, PairsFlag)
    return Original_Sheets


def sheets_recursive(A, df, PairsFlag):   
    for i, record in A.iterrows():
        for cell in record:
            if cell in set(Keys): #name values not 'in' Series (only index values)
                # each key is replaced by the same value
                if PairsFlag == True:
                    A = A.replace(cell, Pairs[0])
                # keys are replaced by corresponding column
                # gets the index location of keypair table: df from A by comparing cell name in df 
                else:
                    i2 = df.loc[df["Name"] == cell].index[0]
                    A = A.replace(cell, df.at[i2, "NewName"])
    return A

New_Sheet_Dict = transform(Original_Sheets, df)



### need to give option for new name
def save_new_files(NewDirect_Path):
    #NewDirect = "/Users/gibson/Desktop/NewRedact/"
    if len(NewDirect_Path) != 0:
        for x in range(0, len(Directory)):
            OldFileNameXLSX = Directory[x]
            #OldFileNameXLSX = File.split("/")[-2]
            OldFileName = OldFileNameXLSX.split(".xlsx")[0]
            NewFileNameXLSX = OldFileName + "_new" + ".xlsx"
            NewFilePath = NewDirect_Path + NewFileNameXLSX
            New_Sheet_Dict[x].to_excel(NewFilePath)
    else:
        for x in range(0, len(Directory)):
            NewFilePath = Directory[x]
            New_Sheet_Dict[x].to_excel(NewFilePath.strip(".xlsx") + "_new.xlsx")
save_new_files(NewDirect_Path)

### need to comment

