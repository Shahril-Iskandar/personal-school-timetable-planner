# !pip install tabula-py
# !pip install "camelot-py[cv]"
# !pip install ghostscript
# !pip install excalibur-py
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

#Updated 15 July 2021

#Select file
file = input("Enter file name: ")

#Allow user to input page to extract
start = input("Input start page: ")
end = input("Input last page: ")
page = f"{start}-{end}"

#Read the pdf and extract the table
tables = camelot.read_pdf(file, pages=page)
print("\nTotal tables extracted: ", tables.n)

#Extract the first page chosen
df = tables[0].df
#Remove first row from second page onwards and append
for t in range(1, tables.n):
  data = tables[t].df
  data = data.drop(labels=0, axis=0)
  df = df.append(data)

#Make first row as the header
new_header = df.iloc[0]
df = df[1:]
df.columns = new_header

#Changing column name without the \n
headerName = ["Course Code", "Title", "Class Type", "Course Type", "Group", "Day", "Time", "Venue", "Remark"]
df.columns = headerName

#Export out first because need to use index_col
df.to_excel("SSM Timetable Planner.xlsx")

#Read the excel file
df = pd.read_excel("SSM Timetable Planner.xlsx", index_col=[0])

#Filling up the empty cells
df.fillna(method="ffill", inplace=True)

#Drop any column that is null, i.e. Remarks
df = df.dropna(axis=1) 

#Replace the \n
for column in df.columns.astype(str):
    df[column] = df[column].str.replace("\n", "")

#If in the 'Day' column has more than 3 characters, remove it
df = df.drop(df[df.Day.str.len() > 3].index)

#Splitting the time into start and end time
df['Start Time'] = df['Time'].str.split('-', expand=True)[0]
df['End Time'] = df['Time'].str.split('-', expand=True)[1]

#Format time to include semicolon
df['Start Time'] = df['Start Time'].str[:2] + ":" + df['Start Time'].str[-2:]
df['End Time'] = df['End Time'].str.strip().str[0:2] + ":"+ df['End Time'].str.strip().str[-2:]

#Delete old time colum
df = df.drop('Time', axis=1)

#Formating the 'Date' column to datetime
week_days= ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
date = str(input('\nEnter the first day of school in dd mm yyyy format, with spaces. \n Example: 9 August 2021 = 09 08 2021 \n Date:  '))
day, month, year = date.split(' ') #10 08 2021
startdate = datetime.date(int(year), int(month), int(day)) #2021-08-10

#From startdate, find the next 6 days.
ls = []
all = []
enddate = startdate + timedelta(days=6) #2021-08-16
diff = enddate - startdate

for i in range(diff.days + 1):
    datee = startdate + datetime.timedelta(i)
    k = datee.strftime("%A") #day name
    d = datee.strftime("%d") #day number
    m = datee.strftime("%m") #month number
    y = datee.strftime("%Y") #year number
    ls.append(k)
    all.append(d + m + y) #to get [10082021, 11082021,...] etc

#Put the date in 10/08/2021 format
newls = []
for i in all:
    i = i[0:2] + "/" + i[2:4] + "/" + i[4:]
    newls.append(i)

#Use the new format so can use the .weekday() function
newday = []
for i in newls:
    newlist = list(map(int, i.split('/')))
    #print(newlist)
    day=datetime.date(newlist[2],newlist[1],newlist[0]).weekday() #1
    d = week_days[day][:3].upper()#TUE
    newday.append(d) #['TUE', 'WED', 'THU', 'FRI', 'SAT', 'SUN', 'MON']

#put in a new dataframe, then merge
list_of_tuples = list(zip(newls, newday))

df2 = pd.DataFrame(list_of_tuples,
                  columns = ['Date', 'Day'])
df2['Date']= pd.to_datetime(df2['Date'], dayfirst=True)

df = pd.merge(df, df2)

#Export out as excel
df.to_excel("SSM Timetable Planner.xlsx")
wb = load_workbook(filename = 'SSM Timetable Planner.xlsx')
ws = wb.active

#Delete the first column which shows the row numbers
ws.delete_cols(idx=1,amount=1)
wb.save("SSM Timetable Planner.xlsx")

print("Success. Please check your file explorer.")
