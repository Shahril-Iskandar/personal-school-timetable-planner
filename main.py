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
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo

file = "SSM Timetable_AY21S1_3.pdf"

#Allow user to input page to extract
start = input("Input start page: ")
end = input("Input last page: ")
page = f"{start}-{end}"

#Read the pdf and extract table
tables = camelot.read_pdf(file, pages=page)
print("Total tables extracted: ", tables.n)

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
table['Course Code'] = table['Course Code'].str.replace("\n", " ")
table['Title'] = table['Title'].str.replace("\n", " ")
table['Course Type'] = table['Course Type'].str.replace("\n", " ")
table['Day'] = table['Day'].str.replace("\n", " ")
table['Time'] = table['Time'].str.replace("\n", " ")
table['Venue'] = table['Venue'].str.replace("\n", " ")

#Dropping sport psychology for ease
table = table.drop(labels=7, axis=0)

#Separate start and end time into new columns and delete old
table['Start Time'] = table['Time'].str.split('-', expand=True)[0]
table['End Time'] = table['Time'].str.split('-', expand=True)[1]
table = table.drop('Time', axis=1)

#Format time to include semicolon
table['Start Time'] = table['Start Time'].str[:2] + ":" + table['Start Time'].str[-2:]
table['End Time'] = table['End Time'].str.strip().str[0:2] + ":"+ table['End Time'].str.strip().str[-2:]

#table['Start Time'] = pd.to_datetime(table['Start Time'], format= '%H:%M').dt.time
#table['End Time'] = pd.to_datetime(table['End Time'], format= '%H:%M').dt.time

#table.info()
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
