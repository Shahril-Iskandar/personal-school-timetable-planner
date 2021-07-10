# !pip install tabula-py
# !pip install "camelot-py[cv]"
# !pip install ghostscript
# !pip install excalibur-py
# !apt install ghostscript python3-tk
# !pip install xlsxwriter
import tabula
import camelot
import pandas as pd
import numpy as np
import openpyxl
import datetime
from datetime import date
from datetime import timedelta
import calendar
import pandas as pd
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo

file = input("Enter file name: ")

#Allow user to input page to extract
start = input("Input start page: ")
end = input("Input last page: ")
page = f"{start}-{end}"

#Read the pdf and extract table
tables = camelot.read_pdf(file, pages=page)
print("\nTotal tables extracted: ", tables.n)

#Remove first row from second page onwards and append
final = tables[0].df
for t in range(1, tables.n):
  data = tables[t].df
  data = data.drop(labels=0, axis=0)
  final = final.append(data)

#display(final)

#Changing column name
headerName = ["Course Code", "Title", "Class Type", "Course Type", "Group", "Day", "Time", "Venue", "Remark"]
final.columns = headerName

#Dropping first row
final = final.drop(labels=0,axis=0)
#display(final)

#Export out
final.to_excel("Table.xlsx")

table = pd.read_excel("Table.xlsx")
table = table.iloc[:,1:] #Remove first column
table = table.replace('', np.nan).ffill() #Fill up the missing blank spaces

table = table.replace('\n', ' ')

#display(table)

table = pd.read_excel("Table.xlsx", dtype=str)
table = table.iloc[:,1:] #Remove first column
table = table.replace('', np.nan).ffill() #Fill up the missing blank spaces

#for column in table.columns:
#  table[column] = table[column].str.replace("\n", " ")
table['Course Code'] = table['Course Code'].str.replace("\n", "")
table['Title'] = table['Title'].str.replace("\n", " ")
table['Course Type'] = table['Course Type'].str.replace("\n", "")
table['Day'] = table['Day'].str.replace("\n", "")
table['Time'] = table['Time'].str.replace("\n", "")
table['Venue'] = table['Venue'].str.replace("\n", "")
table

#Dropping sport psychology for ease
table = table.drop(labels=7, axis=0)

#Separate start and end time into new columns and delete old
table['Start Time'] = table['Time'].str.split('-', expand=True)[0]
table['End Time'] = table['Time'].str.split('-', expand=True)[1]
table = table.drop('Time', axis=1)

#Format time to include semicolon
table['Start Time'] = table['Start Time'].str[:2] + ":" + table['Start Time'].str[-2:]
table['End Time'] = table['End Time'].str.strip().str[0:2] + ":"+ table['End Time'].str.strip().str[-2:]

table = table.dropna(axis=1) #Drop any column that is null, i.e. Remarks
table['Start Time'] = pd.to_datetime(table['Start Time']).dt.time
table['End Time'] = pd.to_datetime(table['End Time']).dt.time

week_days= ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
# l=list(map(int, input("Enter date \n eg: 05 05 2019 \n").split(' ')))
# print("l= ", l) #[10, 8, 2021]
# day=datetime.date(l[2],l[1],l[0]).weekday()
# print(day) #1
# print(week_days[day][:3].upper())#TUE

date = str(input('\nEnter the first day of school in dd mm yyyy format, with spaces. \n Example: 9 August 2021 = 09 08 2019 \n Date:  '))
day, month, year = date.split(' ')
#print(day, month, year) #10 08 2021
startdate = datetime.date(int(year), int(month), int(day)) #2021-08-10
#print(startdate.strftime("%a").upper()) #Tue
#test = [string for string in week_days if startdate.strftime("%a") in string] # ['Tuesday']

ls = []
all = []
enddate = startdate + timedelta(days=6)
print(enddate) #2021-08-16
diff = enddate - startdate

for i in range(diff.days + 1):
    datee = startdate + datetime.timedelta(i)
    k = datee.strftime("%A") #day name
    d = datee.strftime("%d") #day number
    m = datee.strftime("%m") #month number
    y = datee.strftime("%Y") #year number
    ls.append(k)
    all.append(d + m + y)
    #print("k = ", k)
#print("ls = ", ls)
#print("all = ", all) #get [10082021, 11082021,...] etc

newls = []
for i in all:
    i = i[0:2] + "/" + i[2:4] + "/" + i[4:]
    newls.append(i)
    #print(i)
#print("newls = ", newls)

newday = []
for i in newls:
    newlist = list(map(int, i.split('/')))
    print(newlist)
    day=datetime.date(newlist[2],newlist[1],newlist[0]).weekday()
    #print(day) #1
    d = week_days[day][:3].upper()#TUE
    newday.append(d)
#print("newday = ", newday) #['TUE', 'WED', 'THU', 'FRI', 'SAT', 'SUN', 'MON']

list_of_tuples = list(zip(newls, newday))

df = pd.DataFrame(list_of_tuples,
                  columns = ['Date', 'Day'])
df['Date']= pd.to_datetime(df['Date'], dayfirst=True)

table = pd.merge(table, df)
#display(table)

#Export out as excel
table.to_excel("SSM Timetable Planner.xlsx")

wb = load_workbook(filename = 'SSM Timetable Planner.xlsx')
ws = wb.active

ws.delete_cols(idx=1,amount=1)
#for cell in ws['A']:
#    print(cell.value)

#tab = Table(displayName="Table1", ref="A:J")
#style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False,
#                       showLastColumn=False, showRowStripes=True, showColumnStripes=True)
#tab.tableStyleInfo = style
#ws.add_table(tab)
wb.save("SSM Timetable Planner.xlsx")

print("Done. Please check files.")
