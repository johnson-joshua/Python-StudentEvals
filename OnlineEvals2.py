# -*- coding: utf-8 -*-
"""
Created on Fri Jun  8 10:22:17 2018

@author: johnsonjm
"""

import pandas as pd;
import numpy as numpy;

FILE_PATH = "E:\\Stuff\\PythonTest\\BIO-100-H23 - Introductory Biology 12341.201710 - Fall 2017.xlsx";
xl = pd.ExcelFile(FILE_PATH);
writer = pd.ExcelWriter('testing.xlsx', engine='xlsxwriter')
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
        responses['resp'+str(count)] = []
        count+=1
    #print("Responses")
    #print(responses)
        
#Example: addToArray(1, 'Strongly Agree', 4)
def addToArray(qNum, sa, a, d, sd, na):
    #print('qNum = ' +str(int(qNum)))
    #print('Ans = ' +str(answer))
    #print('NumResp = ' +str(numResponses))
    #print("sa=" +str(sa))
    #print("a=" +str(a))
    #print("d=" +str(d))
    #print("sd=" +str(sd))
    #print("na=" +str(na))
#    if answer == 'Strongly Agree' or answer == 'Very Satisfied' or row['Answer'] == 'A':
#        intAnswer = 1
#        answer = 'Strongly Agree' 
#    elif answer == 'Agree' or answer == 'Satisfied' or row['Answer'] == 'B':
#        intAnswer = 2
#        answer = 'Agree'
#    elif answer == 'Disagree' or answer == 'Dissatisfied' or row['Answer'] == 'C':
#        intAnswer = 3
#        answer = 'Disagree'
#    elif answer == 'Strongly Disagree' or answer == 'Very Dissatisfied' or row['Answer'] == 'D':
#        intAnswer = 4
#        answer = 'Strongly Disagree'
#    elif answer == 'Not Applicable' or row['Answer'] == 'F':
#        intAnswer = 5
#        answer = 'Not Applicable'
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
        #responeses['resp'+str(count)].append(int(2))
    #print(responses)
        
        
        
        
        
for index, row in df.iterrows():
    if(row['Q #'] < 38):
           curResp = row['# Responses']
           if numpy.isnan(curResp):
               break
           else:
               totalResp += curResp
numStudents = totalResp/37
makeArrays(int(numStudents))

for index, row in df.iterrows():
    if(row['Q #'] < 38):
        #print(row['Q #'])
        if(row['Answer'] == 'Strongly Agree' or row['Answer'] == 'Very Satisfied' or row['Answer'] == 'A'):
            sa = row['# Responses']
            #addToArray(row['Q #'], row['Answer'], sa)
        elif(row['Answer'] == 'Agree' or row['Answer'] == 'Satisfied' or row['Answer'] == 'B'):
            a = row['# Responses']
            #addToArray(row['Q #'], row['Answer'], a)
        elif(row['Answer'] == 'Disagree' or row['Answer'] == 'Dissatisfied' or row['Answer'] == 'C'):
            d = row['# Responses']
            #addToArray(row['Q #'], row['Answer'], d)
        elif(row['Answer'] == 'Strongly Disagree' or row['Answer'] == 'Very Dissatisfied' or row['Answer'] == 'D'):
            sd = row['# Responses']
            #addToArray(row['Q #'], row['Answer'], sd)
        elif(row['Answer'] == 'Not Applicable' or row['Answer'] == 'F'):
            na = row['# Responses']
            #addToArray(row['Q #'], row['Answer'], na)
            addToArray(row['Q #'], sa, a, d, sd, na)
 
print("Responses Final:")                  
print(responses)