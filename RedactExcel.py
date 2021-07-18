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
  

Directory = os.listdir(input("Enter file directory path"))
NewDirect = input("enter the new directory for your files (include / at end), or leave blank to keep the same: ")


#Directory = os.listdir('/Users/gibson/Desktop/projects/redact/')
Directory.remove(".DS_Store")
Directory.remove("RedactExcel-copy.py")
Directory.remove("pair.xlsx")
Directory.remove("RedactExcelNew.py")


Original_Sheets = {}
i = -1
for items in Directory:
    print(items)
    i += 1
    Original_Sheets[i] = pd.read_excel(items, engine = "openpyxl")


def Create_Key_Pairs():
    redact = {}
    Sheets =[]
    keypair = []
    Keys=""
    Pairs=""
    PairsFlag = False
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
def save_new_files(NewDirect):
    #NewDirect = "/Users/gibson/Desktop/NewRedact/"
    if len(NewDirect) != 0:
        for x in range(0, len(Directory)):
            OldFileNameXLSX = Directory[x]
            #OldFileNameXLSX = File.split("/")[-2]
            OldFileName = OldFileNameXLSX.split(".xlsx")[0]
            NewFileNameXLSX = OldFileName + "_New" + ".xlsx"
            NewFilePath = NewDirect + NewFileNameXLSX
            New_Sheet_Dict[x].to_excel(NewFilePath)
    else:
        for x in range(0, len(Directory)):
            NewFilePath = Directory[x]
            New_Sheet_Dict[x].to_excel(NewFilePath.strip(".xlsx") + "_new.xlsx")
            
save_new_files(NewDirect)

### need to comment

