from xlsxwriter.utility import xl_rowcol_to_cell as getCell
import xlsxwriter as write
import csv
import statistics
import datetime
import glob
import os
import re



def calculate(sheet,lastrow):
    """Add all the calculated values to the sheet"""
    r=lastrow+1
    avgFormat = workbook.add_format()
    avgFormat.set_bold()
    sheet.write(r, 0, "AVERAGE", avgFormat)
    # Loop through each sheet and add in calculations
    for x in range (2, col-1):
        formulaAvg = '=IF(TYPE(' + getCell(1, x) + ')=1,AVERAGE(' + getCell(1, x) + ':' + getCell(r-1, x) + '), "")'
        sheet.write(r, x, formulaAvg, avgFormat)



list_of_files = glob.glob(r'/home/bobk/workspace/gaelhawks_scouting/*.csv')
latestFile = max(list_of_files, key=os.path.getctime)
latestFileName = os.path.basename(max(list_of_files, key=os.path.getctime))
date = datetime.date.today()
now = datetime.datetime.time(datetime.datetime.now())
timestamp = re.sub(r':','',str(now))
timestamp = re.sub(r'\..+','',timestamp)
timestamp = str(date) + "_" + timestamp
excelFile = re.sub(r'RawExport-.+$',timestamp,latestFileName) + ".xlsx"
print(excelFile)

workbook = write.Workbook('C:\\Users\\Jules\\Desktop\\' + excelFile, {'strings_to_numbers':True})
contents = workbook.add_worksheet("Teams")
averages = workbook.add_worksheet("Avg")
predicter = workbook.add_worksheet("Predicter")

try:
    csv_file = open(latestFile)
    csv_reader = csv.DictReader(csv_file)

    teams = {}
    r = 1

    # for each team, create a sheet and populate
    for row in csv_reader:
        teamNum = (row["Team Number"])
        # when we find a new team number create a new sheet
        if (teamNum not in teams.keys()):
            # first finish up the previous sheet
            # ** Note python's crappy way to test if a variable is defined
            try:      # test if variable is defined
                prevteam
            except:   # it is not defined, don't do anything
                pass
            else:     # it is defined, do this stuff
                calculate(teams[prevteam],r)
            prevteam=teamNum
            # now create new sheet and initialize
            teams[row["Team Number"]] = workbook.add_worksheet(row["Team Number"])
            r = 1

        col=0
        for c, (k, v) in enumerate(row.items()):
            # if the key starts with a space, ignore it, it is a header
            if re.match('^ ',k): 
                continue
            if re.match('Team Number',k): 
                continue
            # Add header in top row
            teams[row["Team Number"]].write(0, col, k)
            # Add values to rows
            teams[row["Team Number"]].write(r, col, v)
            col=col+1
        r+=1

except Exception as ee:
    print(ee)

finally:
    csv_file.close()
    
with open (latestFile) as csv_file:
    rawData = workbook.add_worksheet("Raw Data")
    reader = csv.reader(csv_file)
    for r, rows in enumerate(reader):
        for c, col in enumerate(rows):
            rawData.write(r, c, col)


workbook.close()
