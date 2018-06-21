# -*- coding: utf-8 -*-
"""
Created on Fri Jun  8 10:22:17 2018

@author: johnsonjm
"""

import pandas as pd
import numpy as numpy
import os

writer = pd.ExcelWriter('testing.xlsx', engine='xlsxwriter')
path = os.path.expanduser("~/Documents/PythonConverterFolder")
dfExport = pd.DataFrame() #Empty DataFrame to add each class to. This gets exported to Excel
dfComments = pd.DataFrame() #Empty DataFrame to add comments for each class too. This gets exported to the same file as dfExport, but a different worksheet
files = os.listdir(path)

#================================================================================================================
#Excel files to convert with this program MUST be in Documents folder in a folder called "PythonConverterFolder"
#================================================================================================================

for file in files:
    fullName = file.split("-")
    subject = fullName[0]
    classNum = fullName[1]
    section = fullName[2]
    print('Processing: ' +subject, classNum, section)
    FILE_PATH = path+ "/" +file
    xl = pd.ExcelFile(FILE_PATH)
    df = xl.parse('Summary Report')
    curResp = 0
    totalResp = 0
    numStudents = 0
    commentsIterator = 1
    sa = 0
    a = 0
    d = 0
    sd = 0
    na = 0
    
    #responses dictionary will hold a list for each student that took survey (Q 1-37)
    responses = {}
    #comments dictionary will hold a list for each student that took the survey (comments)
    comments = {}
    
    #Makes a list for each response
    def makeArrays(numResponses):
        count = int(1)
        for response in range(int(numResponses)):
            responses['resp'+str(count)] = [subject,classNum,section] #adds class data to the first 3 columns of each list
            count+=1
    
    #Makes a list for each response
    def makeCommentArrays(numResponses):
        count = int(1)
        for response in range(int(numResponses)):
            comments['com'+str(count)] = [subject, classNum, section]
            count+=1
    
    #Adds data to each list for Q1-Q34 & Q36-END
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
            
    #Adds data to each list for Q35 (comment question)
    def addCommentsToArray(count, qNum, qText, comment):
        comments['com'+str(count)].append(qNum)
        comments['com'+str(count)].append(qText)
        comments['com'+str(count)].append(comment)

    #Q1 MUST be required. This logic will loop through Q1 and add up the number of responses for "Strongly Agree", "Agree", etc
    for index, row in df.iterrows():
        if(row['Q #'] == 1.0):
               curResp = row['# Responses']
               if numpy.isnan(curResp):
                   break
               else:
                   totalResp += curResp
    numStudents = totalResp
    makeArrays(int(numStudents))
    makeCommentArrays(int(numStudents))
    
    #Iterates through the original Excel report from D2L. Loops through Q1-Q34 & Q36-END and stores the number of responses for strongly agree, agree, etc.
    #Once the loops hits the 'Not Applicable' number of responses, it calls the method to add the data for that question into a list
    #
    for index, row in df.iterrows():
        if(row['Q #'] <= 34.0 or row['Q #'] >= 36.0):
            if(row['Answer'] == 'Strongly Agree' or row['Answer'] == 'Very Satisfied' or row['Answer'] == 'A'):
                sa = row['# Responses']
            if(row['Answer'] == 'Agree' or row['Answer'] == 'Satisfied' or row['Answer'] == 'B'):
                a = row['# Responses']
            if(row['Answer'] == 'Disagree' or row['Answer'] == 'Dissatisfied' or row['Answer'] == 'C'):
                d = row['# Responses']
            if(row['Answer'] == 'Strongly Disagree' or row['Answer'] == 'Very Dissatisfied' or row['Answer'] == 'D'):
                sd = row['# Responses']
            if(row['Answer'] == 'Not Applicable' or row['Answer'] == 'F'):
                na = row['# Responses']
                addToArray(row['Q #'], sa, a, d, sd, na) #AddToArray is called once per question (1-34 && 36-END)
        elif(row['Q #'] == 35.0):
           addCommentsToArray(commentsIterator, row['Q #'], row['Q Text'], row['Answer'])
           commentsIterator+=1
                 
    
    dfTemp = pd.DataFrame.from_dict(responses, orient='index', 
                                    columns=['A','B','C','1','2','3','4','5','6','7','8','9','10','11','12','13','14','15','16','17','18','19','20','21','22','23','24','25','26','27','28','29','30','31','32','33','34','36','37','38'])
    dfTemp2 = pd.DataFrame.from_dict(comments, orient='index')
    dfExport = dfExport.append(dfTemp, ignore_index=True)
    dfComments = dfComments.append(dfTemp2, ignore_index=True)
dfExport.to_excel(writer, sheet_name="Results")
dfComments.to_excel(writer, sheet_name="Comments") #New worksheet holds all comments
writer.save()
