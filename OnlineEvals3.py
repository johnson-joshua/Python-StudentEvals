# -*- coding: utf-8 -*-
"""
Created on Fri Jun  8 10:22:17 2018

@author: johnsonjm
"""

import pandas as pd;
import numpy as numpy;
import os;

writer = pd.ExcelWriter('testing.xlsx', engine='xlsxwriter')
path = os.path.expanduser("~/Documents/PythonConverterFolder")
dfExport = pd.DataFrame()
files = os.listdir(path)

#================================================================================================================
#Excel files to convert with this program MUST be in Documents folder in a folder called "PythonConverterFolder"
#================================================================================================================

for file in files:
    fullName = file.split("-")
    subject = fullName[0]
    classNum = fullName[1]
    section = fullName[2]
    print(subject, classNum, section)
    FILE_PATH = path+ "/" +file
    xl = pd.ExcelFile(FILE_PATH);
    df = xl.parse('Summary Report');
    curResp = 0
    totalResp = 0
    numStudents = 0
    sa = 0
    a = 0
    d = 0
    sd = 0
    na = 0
    
    #responses dictionary will hold a list for each student that took survey
    responses = {}
    
    #Makes a list for each response
    def makeArrays(numResponses):
        count = int(1)
        for response in range(int(numResponses)):
            responses['resp'+str(count)] = [subject,classNum,section] #adds class data to the first 3 columns of each list
            count+=1
    
    
    def addToArray(qNum, sa, a, d, sd, na):
        count = int(1)
        for response in range(int(sa)):
            responses['resp'+str(count)].append(int(1))
            count+=1
        for response in range(int(a)):
            responses['resp'+str(count)].append(int(2))
            count+=1
        for response in range(int(d)):
            responses['resp'+str(count)].append(int(3))
            count+=1
        for response in range(int(sd)):
            responses['resp'+str(count)].append(int(4))
            count+=1
        for response in range(int(na)):
            responses['resp'+str(count)].append(int(5))
            count+=1
    
    #The logic for finding the number of responses in a given file. Add up responses for questions 1-37 and divide by 37
    #TODO: Change the logic to loop through Question 1 and get the number of responses. Question 1 will be required within D2L
    for index, row in df.iterrows():
        if(row['Q #'] < 38):
               curResp = row['# Responses']
               if numpy.isnan(curResp):
                   break
               else:
                   totalResp += curResp
    numStudents = totalResp/37
    makeArrays(int(numStudents))
    
    #Iterates through the original Excel report from D2L. Loops through questions 1-37 and stores the number of responses for strongly agree, agree, etc.
    #Once the loops hits the 'Not Applicable' number of responses, it calls the method to add the data for that question into a list
    #TODO: Loop through the rest of spreadsheet (38 - END) and store comments in their own worksheet
    for index, row in df.iterrows():
        if(row['Q #'] < 38):
            if(row['Answer'] == 'Strongly Agree' or row['Answer'] == 'Very Satisfied' or row['Answer'] == 'A'):
                sa = row['# Responses']
            elif(row['Answer'] == 'Agree' or row['Answer'] == 'Satisfied' or row['Answer'] == 'B'):
                a = row['# Responses']
            elif(row['Answer'] == 'Disagree' or row['Answer'] == 'Dissatisfied' or row['Answer'] == 'C'):
                d = row['# Responses']
            elif(row['Answer'] == 'Strongly Disagree' or row['Answer'] == 'Very Dissatisfied' or row['Answer'] == 'D'):
                sd = row['# Responses']
            elif(row['Answer'] == 'Not Applicable' or row['Answer'] == 'F'):
                na = row['# Responses']
                #AddToArray is called once per question (1-37)
                addToArray(row['Q #'], sa, a, d, sd, na)
    
    dfTemp = pd.DataFrame.from_dict(responses, orient='index')
    print(dfTemp)
    dfExport = dfExport.append(dfTemp, ignore_index=True)
    print(dfExport)
dfExport.to_excel(writer, sheet_name="Results")
writer.save()
