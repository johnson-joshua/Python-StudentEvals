import pandas as pd;
import numpy as numpy;

FILE_PATH = "E:\\Stuff\\PythonTest\\BIO-100-H23 - Introductory Biology 12341.201710 - Fall 2017.xlsx";
xl = pd.ExcelFile(FILE_PATH);
writer = pd.ExcelWriter('testing.xlsx', engine='xlsxwriter')
df = xl.parse('Summary Report');
#df2 = pd.DataFrame(columns=["a", "b","c","faculty code","semester", "name", "division","1","2","3","4","5","6","7","8","9","10","11","12","13","14","15","16","17","18","19","20","21","22","23","24","25","26","27","28","29","30","31","32","33","34","35","36","37"])
df2 = pd.DataFrame(index=range(1,1000), columns=["a", "b", "c", "faculty code", "semester", "name", "division"]);
nArray = numpy.zeros(shape=df2.shape)

resp = 0
totalResp = 0
numStudents = 0
sa = 0
a = 0
d = 0
sd = 0
na = 0
exportRow = 0

#make a dictionary for each question
#key: Strongly Agree, Agree, etc
myList1 = {}
myList2 = {}
myList3 = {}
myList4 = {}
myList5 = {}
myList6 = {}
myList7 = {}
myList8 = {}
myList9 = {}
myList10 = {}
myList11 = {}
myList12 = {}
myList13 = {}
myList14 = {}
myList15 = {}
myList16 = {}
myList17 = {}
myList18 = {}
myList19 = {}
myList20 = {}
myList21 = {}
myList22 = {}
myList23 = {}
myList24 = {}
myList25 = {}
myList26 = {}
myList27 = {}
myList28 = {}
myList29 = {}
myList30 = {}
myList31 = {}
myList32 = {}
myList33 = {}
myList34 = {}
myList35 = {}
myList36 = {}
myList37 = {}

def addToArray(qNum, answer, numResponses):
    print('qNum = ' +str(int(qNum)))
    #print('Ans = ' +str(answer))
    #print('NumResp = ' +str(numResponses))
    
    values = []
    #values = numpy.zeros
    if answer == 'Strongly Agree' or answer == 'Very Satisfied' or row['Answer'] == 'A':
        intAnswer = 1
        answer = 'Strongly Agree' 
    elif answer == 'Agree' or answer == 'Satisfied' or row['Answer'] == 'B':
        intAnswer = 2
        answer = 'Agree'
    elif answer == 'Disagree' or answer == 'Dissatisfied' or row['Answer'] == 'C':
        intAnswer = 3
        answer = 'Disagree'
    elif answer == 'Strongly Disagree' or answer == 'Very Dissatisfied' or row['Answer'] == 'D':
        intAnswer = 4
        answer = 'Strongly Disagree'
    elif answer == 'Not Applicable' or row['Answer'] == 'F':
        intAnswer = 5
        answer = 'Not Applicable'
    for response in range(int(numResponses)):
        values.append(intAnswer) #add a number of answers to a list    
        #values = numpy.append(values, intAnswer)
    if int(qNum) == 1:
        myList1[answer] = values
        print(myList1)
    elif int(qNum) == 2:
        myList2[answer] = values
        print(myList2)
    elif int(qNum) == 3:
        myList3[answer] = values
        print(myList3)
    elif int(qNum) == 4:
        myList4[answer] = values
        print(myList4)
    elif int(qNum) == 5:
        myList5[answer] = values
        print(myList5)
    elif int(qNum) == 6:
        myList6[answer] = values
        print(myList6)
    elif int(qNum) == 7:
        myList7[answer] = values
        print(myList7)
    elif int(qNum) == 8:
        myList8[answer] = values
        print(myList8)
    elif int(qNum) == 9:
        myList9[answer] = values
        print(myList9)
    elif int(qNum) == 10:
        myList10[answer] = values
        print(myList10)
    elif int(qNum) == 11:
        myList11[answer] = values
        print(myList11)
    elif int(qNum) == 12:
        myList12[answer] = values
        print(myList12)
    elif int(qNum) == 13:
        myList13[answer] = values
        print(myList13)
    elif int(qNum) == 14:
        myList14[answer] = values
        print(myList14)
    elif int(qNum) == 15:
        myList15[answer] = values
        print(myList15)
    elif int(qNum) == 16:
        myList16[answer] = values
        print(myList16)
    elif int(qNum) == 17:
        myList17[answer] = values
        print(myList17)
    elif int(qNum) == 18:
        myList18[answer] = values
        print(myList18)
    elif int(qNum) == 19:
        myList19[answer] = values
        print(myList19)
    elif int(qNum) == 20:
        myList20[answer] = values
        print(myList20)
    elif int(qNum) == 21:
        myList21[answer] = values
        print(myList21)
    elif int(qNum) == 22:
        myList22[answer] = values
        print(myList22)
    elif int(qNum) == 23:
        myList23[answer] = values
        print(myList23)
    elif int(qNum) == 24:
        myList24[answer] = values
        print(myList24)
    elif int(qNum) == 25:
        myList25[answer] = values
        print(myList25)
    elif int(qNum) == 26:
        myList26[answer] = values
        print(myList26)
    elif int(qNum) == 27:
        myList27[answer] = values
        print(myList27)
    elif int(qNum) == 28:
        myList28[answer] = values
        print(myList28)
    elif int(qNum) == 29:
        myList29[answer] = values
        print(myList29)
    elif int(qNum) == 30:
        myList30[answer] = values
        print(myList30)
    elif int(qNum) == 31:
        myList31[answer] = values
        print(myList31)
    elif int(qNum) == 32:
        myList32[answer] = values
        print(myList32)
    elif int(qNum) == 33:
        myList33[answer] = values
        print(myList33)
    elif int(qNum) == 34:
        myList34[answer] = values
        print(myList34)
    elif int(qNum) == 35:
        myList35[answer] = values
        print(myList35)
    elif int(qNum) == 36:
        myList36[answer] = values
        print(myList36)
    elif int(qNum) == 37:
        myList37[answer] = values
        print(myList37)
    
for index, row in df.iterrows():
    if(row['Q #'] < 38):
        #print(row['Q #'])
        if(row['Answer'] == 'Strongly Agree' or row['Answer'] == 'Very Satisfied' or row['Answer'] == 'A'):
            sa = row['# Responses']
            addToArray(row['Q #'], row['Answer'], sa)
        elif(row['Answer'] == 'Agree' or row['Answer'] == 'Satisfied' or row['Answer'] == 'B'):
            a = row['# Responses']
            addToArray(row['Q #'], row['Answer'], a)
        elif(row['Answer'] == 'Disagree' or row['Answer'] == 'Dissatisfied' or row['Answer'] == 'C'):
            d = row['# Responses']
            addToArray(row['Q #'], row['Answer'], d)
        elif(row['Answer'] == 'Strongly Disagree' or row['Answer'] == 'Very Dissatisfied' or row['Answer'] == 'D'):
            sd = row['# Responses']
            addToArray(row['Q #'], row['Answer'], sd)
        elif(row['Answer'] == 'Not Applicable' or row['Answer'] == 'F'):
            na = row['# Responses']
            addToArray(row['Q #'], row['Answer'], na)
        resp = row['# Responses']
        if numpy.isnan(resp):
            break
        else:
            totalResp += resp
            
numStudents = totalResp/37                                   

df3 = pd.DataFrame(columns=["a", "b","c","faculty code","semester", "name", "division","1","2","3","4","5","6","7","8","9","10","11","12","13","14","15","16","17","18","19","20","21","22","23","24","25","26","27","28","29","30","31","32","33","34","35","36","37"])
dictList = [myList1, myList2, myList3, myList4, myList5,myList6, myList7, myList8, myList9, myList10,myList11, myList12, myList13, myList14, myList15,myList16, myList17, myList18, myList19, myList20,myList21, myList22, myList23, myList24, myList25,myList26, myList27, myList28, myList29, myList30,myList31, myList32, myList33, myList34, myList35,myList36, myList37]
for curList in dictList:
    #print(curList)
    if curList != {}:
        sa = curList['Strongly Agree']
        a = curList['Agree']
        d = curList['Disagree']
        sd = curList['Strongly Disagree']
        na = curList['Not Applicable']
        print(sa)
        print(a)
        print(d)
        print(sd)
        print(na)

#print(myList1)
#df3 = pd.DataFrame([myList1, myList2, myList3, myList4, myList5,myList6, myList7, myList8, myList9, myList10,myList11, myList12, myList13, myList14, myList15,myList16, myList17, myList18, myList19, myList20,myList21, myList22, myList23, myList24, myList25,myList26, myList27, myList28, myList29, myList30,myList31, myList32, myList33, myList34, myList35,myList36, myList37])
df3.to_excel(writer, sheet_name='Sheet1')