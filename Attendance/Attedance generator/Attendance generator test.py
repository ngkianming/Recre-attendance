###Test code for blacklist on GDrive

import pandas as pd
import numpy as np
import xlsxwriter
import sys
'''
print('Reminder: save the csv file again after downloading')

## Inputs ##
month = input("Month: ")

num_trng = int(input("Number of trainings: "))

loop = True

tele = 13
while loop:
    tele_col = input("Is the tele handle in column N? [Y/N]: ")
    if tele_col.upper() == 'N':
        tele = ord(input('which column is it in: ').lower()) - 97
        loop = False
    elif tele_col.upper() == 'Y':
        loop = False
    else:
        print('invalid input, please input y/n')
    
loop = True

start = 14

while loop:
    training_col = input("Is the first trining in column O? [Y/N]: ")
    if training_col.upper() == 'N':
        start = ord(input('which column is it in: ').lower()) - 97
        loop = False
    elif training_col.upper() == 'Y':
        loop = False
    else:
        print('invalid input, please input y/n')

question = input('Question: ')
qlength = len(question)


#index for dates
stop = start + num_trng - 1
'''

## Codes ##
'''
#Blacklist#
sheet_url = "https://docs.google.com/spreadsheets/d/1ZHAWpgebGMW0q8GgTLmjXLO7gV9vsm7Lo0K-njnzxnc/edit#gid=0"
url_1 = sheet_url.replace('/edit#gid=', '/export?format=csv&gid=')
blacklist = pd.read_csv(url_1, header = 0, on_bad_lines='skip')
print(blacklist)
'''

atd_sheet = pd.read_excel("../Attendance summary.xlsx", index_col = 0, sheet_name = "Summary (%)",header=2)
atd_sheet = atd_sheet.iloc[:,0:8]
low_pio = atd_sheet.loc[atd_sheet['Grand Total']>=3]
low_pio = low_pio.loc[atd_sheet['Attended']<0.5]
low_pio = low_pio.iloc[:,0:2]
low_pio.columns = ['name','Username']

pd.set_option("display.max_rows", None, "display.max_columns", None)
print(atd_sheet)                      

'''
#Attendance#
#read file name
file = month + " raw.csv" 
df = pd.read_csv(file, header = 2)

#data extraction
dates = []
for i in range(stop - start+1):  #extract dates
    dates.append(df.columns[start + i][qlength:qlength+6])

#data sorting and export
file_export = month + '.xlsx'
writer = pd.ExcelWriter(file_export, engine='xlsxwriter')

namelist = pd.DataFrame()
sign_ups = []
formula = [['Nvr pld', '=COUNTIF(G:G,"I have never played floorball before, or for less than 6 months")'],
           ['Beginner', '=COUNTIF(G:G,"Beginner (6 months - 1 year)")'],
           ['Intermediate', '=COUNTIF(G:G,"Intermediate (1 year - 3 years, participated in friendlies)")'],
           ['Advanced', '=COUNTIF(G:G,"Advanced (> 3 years, represented school for competitions)")']]

for i in range(len(dates)):
    #generate excel sheet for dates
    new_df = df.iloc[:,np.r_[1,2,3,4,tele,i+start, stop +4, stop + 5]]
    new_df.dropna(inplace = True)
    new_df['DateSubmitted']=pd.to_datetime(new_df['DateSubmitted']) 
    sort = new_df.sort_values(by='DateSubmitted')
    sort['Name'] = sort['First Name'] + ' ' + sort['Last Name']
    sort = sort.iloc[:,np.r_[0,8,1,4,5,6,7]]
    sort = pd.merge(sort, blacklist, on=["Username"], how="outer", indicator = True)
    removed = sort.loc[sort["_merge"] == 'both'].drop(["_merge","name"], axis = 1)
    sort = sort.loc[sort["_merge"] == 'left_only'].drop(["_merge","name"], axis = 1)
    sort.to_excel(writer, sheet_name=dates[i], index = False)
    removed.to_excel(writer, sheet_name=dates[i], index = False, header = False, startcol = 0, startrow = len(sort.index)+2)
    #generate data and namelist
    formula_df = pd.DataFrame(formula)
    formula_df.to_excel(writer, sheet_name=dates[i], index = False, header = False, startcol = 14, startrow = 1)
    namelist = pd.concat([namelist,sort.iloc[:,np.r_[1,2]],removed.iloc[:,np.r_[1,2]]])
    sign_ups.append([dates[i],len(sort)])

namelist.drop_duplicates(inplace=True)
namelist.to_excel(writer, sheet_name='Namelist', index = False)

signup_df = pd.DataFrame(sign_ups, columns = ['Dates','Number of sign ups'])
signup_df.to_excel(writer, sheet_name='Namelist',startcol = 4, startrow = 0, index = False)

writer.save()

print('Done! Press enter to exit')
input()
'''
