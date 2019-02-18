import xlrd
import xlsxwriter as write
import csv
import statistics
import datetime
import glob
import os
import re

list_of_files = glob.glob(r'C:\Users\Jules\Dropbox\Scouting 2019\*.csv')
latestFile = max(list_of_files, key=os.path.getctime)
latestFileName = os.path.basename(max(list_of_files, key=os.path.getctime))
date = datetime.date.today()
now = datetime.datetime.time(datetime.datetime.now())
timestamp = re.sub(r':','',str(now))
timestamp = re.sub(r'\..+','',timestamp)
timestamp = str(date) + "_" + timestamp
excelFile = re.sub(r'RawExport-.+$',timestamp,latestFileName) + ".xlsx"
print(excelFile)
workbook = write.Workbook('C:\\Users\\Jules\\Desktop\\' + excelFile)
contents = workbook.add_worksheet("Teams")
averages = workbook.add_worksheet("Avg")
predicter = workbook.add_worksheet("CrystalBall")
try:
    csv_file = open(latestFile)
    csv_reader = csv.DictReader(csv_file)
    teams = {}
    r = 0
    for row in csv_reader:
        teamNum = (row["Team Number"])
        if (teamNum not in teams.keys()):
            teams[row["Team Number"]] = workbook.add_worksheet(row["Team Number"])
            r = 0
        for c, col in enumerate(row):
            teams[row["Team Number"]].write(0, c, col)
        for ci in row.keys():
            for c, col in enumerate(row.items()):                
                teams[row["Team Number"]].write(r+1, c, col[1])
        r+=1
finally:
    csv_file.close()
    
with open (latestFile) as csv_file:
    rawData = workbook.add_worksheet("Raw Data")
    reader = csv.reader(csv_file)
    for r, rows in enumerate(reader):
        for c, col in enumerate(rows):
            rawData.write(r, c, col)


workbook.close()
