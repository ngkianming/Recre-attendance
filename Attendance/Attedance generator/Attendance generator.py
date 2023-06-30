import pandas as pd
import numpy as np
import xlsxwriter
import sys
from tqdm import tqdm

import warnings
warnings.filterwarnings('ignore') #ignore warning

##### Inputs #####
month = input("Month: ")

num_trng = int(input("Number of trainings: "))

tele = ord(input('which column is tele handle in: ').lower()) - 97

start = ord(input('which column is first training in: ').lower()) - 97

question = input('Question: ')
qlength = len(question)

print('Reminder: save the csv file again after downloading')
print('Check the tele handles for any sus looking handles')
input("press enter to generate attendance")


#index for dates
stop = start + num_trng - 1


##### Codes #####

##Filtering list##

#blacklist#
blacklist = pd.read_csv("Blacklist.csv", header = 0)

#low piority#
#namelist
atd_sheet = pd.read_excel("../Attendance summary.xlsx", index_col = 0, sheet_name = "Summary (%)",header=3)
atd_sheet = atd_sheet.iloc[:,0:9]
main_lst = atd_sheet.loc[atd_sheet['Grand Total']>2] #new player defination
#print(main_lst.iloc[:,0:2].to_string())
lst = main_lst.iloc[:,1]

#by high withdrawn (total >4, withdrawn > 80%)
low_pio_attd = main_lst.loc[atd_sheet['Grand Total']>4]
low_attd = low_pio_attd.loc[atd_sheet['Withdrawals']>0.8]
low_attd = low_attd.iloc[:,0:2]
low_attd.columns = ['name','Username']

#by no show (confirmed >2, absent > 50%)
low_pio_noshow = main_lst.loc[atd_sheet['Confirmed']*atd_sheet['Grand Total']>2]
low_noshow = low_pio_noshow.loc[atd_sheet['Absent']/atd_sheet['Confirmed']>0.5]
low_noshow = low_noshow.iloc[:,0:2]
low_noshow.columns = ['name','Username']

#by late (attended >3, late > 2/3)
low_pio_late = main_lst.loc[atd_sheet['Grand Total']*atd_sheet['Attended']>3]
low_late = low_pio_late.loc[atd_sheet['Late']/atd_sheet['Attended']>0.66]
low_late = low_late.iloc[:,0:2]
low_late.columns = ['name','Username']


##Attendance generation##

#read file name#
file = month + " raw.csv" 
df = pd.read_csv(file, header = 2)

#dates extraction#
dates = []
for i in range(stop - start+1):  #extract dates
    date = df.columns[start + i][qlength:qlength+2]
    if len(date)-date.count(" ")==1:
        dates.append(df.columns[start + i][qlength:qlength+5])
    else:
        dates.append(df.columns[start + i][qlength:qlength+6])

#Initialising#
file_export = month + '.xlsx'
writer = pd.ExcelWriter(file_export, engine='xlsxwriter')

namelist = pd.DataFrame()
sign_ups = []
formula = [['Nvr pld', '=COUNTIF(G:G,"I have never played floorball before, or for less than 6 months")'],
           ['Beginner', '=COUNTIF(G:G,"Beginner (6 months - 1 year)")'],
           ['Intermediate', '=COUNTIF(G:G,"Intermediate (1 year - 3 years, participated in friendlies)")'],
           ['Advanced', '=COUNTIF(G:G,"Advanced (> 3 years, represented school for competitions)")']]
#database = pd.read_csv('Database.csv')
#database = pd.DataFrame() #for if database is blank

#filtering function#
def filtering(main_list, filtered_list):
    sort = pd.merge(main_list, filtered_list, on=["Username"], how="outer", indicator = True)
    low = sort.loc[sort["_merge"] == 'both'].drop(["_merge","name"], axis = 1)
    sort = sort.loc[sort["_merge"] == 'left_only'].drop(["_merge","name"], axis = 1)
    return [sort,low]

#Namelist generation#
for i in tqdm(range(len(dates))):
    #generate df for each dates#
    new_df = df.iloc[:,np.r_[1,2,3,4,tele,i+start, stop +4, stop + 5]]
    new_df.dropna(inplace = True)
    new_df['DateSubmitted']=pd.to_datetime(new_df['DateSubmitted']) 
    sort = new_df.sort_values(by='DateSubmitted')
    sort['Name'] = sort['First Name'] + ' ' + sort['Last Name']
    sort = sort.iloc[:,np.r_[0,8,1,4,5,6,7]]
    sort['Username'] = sort['Username'].str.upper()

    #low piority (late)#
    sort = pd.concat(filtering(sort, low_late))

    #low piority (attendance)#
    sort = pd.concat(filtering(sort, low_attd))
    
    #low piority (no show)#
    sort = pd.concat(filtering(sort, low_noshow))

    #low piority (staff members)#
    low_staff = sort[sort["Username"].str.contains("NUSSTF", case=False)]
    low_staff = low_staff.iloc[:,np.r_[1,2]]
    low_staff.columns = ['name','Username']
    sort = pd.concat(filtering(sort, low_staff))

    #blacklist#
    removed = filtering(sort, blacklist)[1]
    sort = filtering(sort, blacklist)[0]

    #new player stats#
    sort["<=2 trainings"] = np.where(sort['Username'].str.upper().isin(lst),False,True)
    
    sort.to_excel(writer, sheet_name=dates[i], index = False)
    removed.to_excel(writer, sheet_name=dates[i], index = False, header = False, startcol = 0, startrow = len(sort.index)+2)

    #generate data and namelist#
    formula_df = pd.DataFrame(formula)
    formula_df.to_excel(writer, sheet_name=dates[i], index = False, header = False, startcol = 14, startrow = 1) #writing namelist to excel file
    namelist = pd.concat([namelist,sort.iloc[:,np.r_[1,2]],removed.iloc[:,np.r_[1,2]]])
    sign_ups.append([dates[i],len(sort)])


#Generate namelist and stats of the month#
namelist.drop_duplicates(inplace=True)
for columns in namelist.columns:
    namelist[columns] = namelist[columns].str.upper() 
namelist.to_excel(writer, sheet_name='Namelist', index = False)

signup_df = pd.DataFrame(sign_ups, columns = ['Dates','Number of sign ups'])
signup_df['Dates']=pd.to_datetime(signup_df['Dates'], format="%d %b") 
signup_df = signup_df.sort_values(by='Dates')
signup_df['Dates']=signup_df['Dates'].dt.strftime('%d %b') 
signup_df.to_excel(writer, sheet_name='Namelist',startcol = 4, startrow = 0, index = False)

writer.save()

#database.drop_duplicates(inplace=True, subset = ['Username'])
#database.to_csv('Database.csv',index = False)

print("\n")
print("High withdrawn")
print(low_attd)
print("\n")
print("High Late")
print(low_late)
print("\n")
print("High Absent")
print(low_noshow)

print("\n")

print('Done! Press enter to exit')
input()

