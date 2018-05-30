import pandas as pd;
import numpy as numpy;

FILE_PATH = "E:\\Stuff\\PythonTest\\BIO-100-H23 - Introductory Biology 12341.201710 - Fall 2017.xlsx";
xl = pd.ExcelFile(FILE_PATH);
writer = pd.ExcelWriter('testing.xlsx', engine='xlsxwriter')
df = xl.parse('Summary Report');
#df2 = pd.DataFrame(columns=["a", "b","c","faculty code","semester", "name", "division","1","2","3","4","5","6","7","8","9","10","11","12","13","14","15","16","17","18","19","20","21","22","23","24","25","26","27","28","29","30","31","32","33","34","35","36","37"])
df2 = pd.DataFrame(index=range(1,1000), columns=["a", "b", "c", "faculty code", "semester", "name", "division"]);
resp = 0
totalResp = 0
numStudents = 0
sa = 0
a = 0
d = 0
sd = 0
na = 0
exportRow = 0

#for index, row in df.iterrows():
#    print(index, row['Q #'], row['Answer'], row['# Responses'])

def addToDataFrame(qNum, answer, numResponses):
    print('qNum = ' +str(qNum))
    print('Ans = ' +str(answer))
    print('NumResp = ' +str(numResponses))
    if answer == 'Strongly Agree':
        intAnswer = 1
    elif answer == 'Agree':
        intAnswer = 2
    elif answer == 'Disagree':
        intAnswer = 3
    elif answer == 'Strongly Disagree':
        intAnswer = 4
    elif answer == 'Not Applicable':
        intAnswer = 5
        
    global exportRow
    count = 1
    for response in range(int(numResponses)):
        #print('Response = ' +str(response))
        #print('exportRow = ' +str(exportRow))
        if count > int(numResponses):
            count = 1
        else:
            count += 1
            print('Count=' +str(count))
        
        df2.at[count, int(qNum)] = intAnswer
        print('Answer = ' +str(intAnswer))

        #exportRow += 1
    
for index, row in df.iterrows():   
    if(row['Q #'] < 38):
        #print('Q #= ' +str(row['Q #']))
        if(row['Answer'] == 'Strongly Agree'):
            sa = row['# Responses']
#            print('SA: ' +str(sa))
            addToDataFrame(row['Q #'], row['Answer'], sa)
        elif(row['Answer'] == 'Agree'):
            a = row['# Responses']
            addToDataFrame(row['Q #'], row['Answer'], a)
#            print('A: ' +str(a))
        elif(row['Answer'] == 'Disagree'):
            d = row['# Responses']
            addToDataFrame(row['Q #'], row['Answer'], d)
#            print('D: ' +str(d))
        elif(row['Answer'] == 'Strongly Disagree'):
            sd = row['# Responses']
            addToDataFrame(row['Q #'], row['Answer'], sd)
#            print('SD: ' +str(sd))
        elif(row['Answer'] == 'Not Applicable'):
            na = row['# Responses']
            addToDataFrame(row['Q #'], row['Answer'], na)
#            print('NA: ' +str(na))
        
        resp = row['# Responses']
        if numpy.isnan(resp):
            break
        else:
            totalResp += resp
            #print('TotalResp: ' +str(totalResp))
            #print('Number Responses: ' +str(totalResp))
numStudents = totalResp/37                                   
#print('Number of Students: '+str(numStudents))


#get num responses for each value (SA, A, D, SD, NA)

#for stu in range(int(numStudents)):
#    df2 = df2.append({
#        "a": "BIO",
#        "b": "100",
#        "c": "H23",
#        "1": "John",
#        "2":  "Johny"
#    }, ignore_index=True)

df2.to_excel(writer, sheet_name='Sheet1')
#print(df2)

#answers = {'Strongly Agree':1, 'Agree':2, 'Disagree':3, 'Strongly Disagree':4, 'Not Applicable':5}
#for k, v in answers.items():
#    print(k, v)
#print(answers['Strongly Agree'])