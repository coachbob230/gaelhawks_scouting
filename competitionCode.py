from xlsxwriter.utility import xl_rowcol_to_cell as getCell
import xlsxwriter as write
import csv
import datetime
import glob
import os
import re



def calculate(sheet,lastrow):
    """ Add all the calculated values to the sheet """

    r=lastrow+2
    avgFormat = workbook.add_format()
    avgFormat.set_bold()
    sheet.write(r, 0, "AVERAGE", avgFormat)
    # Loop through each sheet and add in calculations
    for x in range (2, col-1):
        formulaAvg = '=IF(TYPE(' + getCell(1, x) + ')=1,AVERAGE(' + getCell(1, x) + ':' + getCell(r-1, x) + '), "")'
        sheet.write(r, x, formulaAvg, avgFormat)

    # next line of calculations    
    r = r + 1
    #sheet.write(r, 2, "Total Items:")
    #sheet.write(r, 3, "=SUM()")



def predicterPage(sheet):
    """ Create the predicter page """
    pass



def teamPage(sheet,teamNum):
    """ Create the team page """

    # TODO: get data from TBA
    pass



def getTBATeamData(teamNum):
    """ Get The Blue Alliance team data iff we have an internet connection
    Store the data in a file so that we can use it later and add it to the team sheets"""
    pass



def readTBATeamData(teamNum):
    """ Read the team data from the file and put it in the team's tab """
    pass



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

# Get event name from the file name
(_,event,_)=latestFileName.split(r'_',2)
event=re.sub(r' Event','',event)
#print("Event: " + event)

# Specify the order that the columns should be in in the speadsheet
# TODO: put this in a config file
columnOrder={
    # Column Name:           Col Num
    "Team Number":                   0,
    "Match Number":   	            11,
    "SANDSTORM":                    20,
    "Starting Level":               21,
    "Start With Hatch or Cargo?":   22,
    "Leave Hab?":                   23,
    "Cargo Placed":                 24,
    "Hatch Panels Placed":          25,
    "TELE-OP":                      30,
    "Hatch on Cargo Ship":          31,
    "Hatch on Rocket LOW":          32,
    "Hatch on Rocket MID":          33,
    "Hatch on Rocket TOP":          34,
    "Cargo in Cargo Ship":          35,
    "Cargo in Rocket LOW":          36,
    "Cargo on Rocket MID":          37,
    "Cargo on Rocket TOP":          38,
    "END GAME":                     40,
    "End Hab Level":                41,
    "Did They Get Lifted?":         42,
    "Did They Lift Another Robot?": 43,
    "Comments":                     50,
    "Match Comment":                51
}


def getOrder(key):
    """ return the value which puts the columns in the correct order """

    try:
        k=key[0].lstrip()
        columnOrder[k]
    except Exception:
        return 999
    else:
        return columnOrder[k]




workbook = write.Workbook('C:\\Users\\Jules\\Desktop\\' + excelFile, {'strings_to_numbers':True})
contents = workbook.add_worksheet("Teams")
averages = workbook.add_worksheet("Avg")
predicter = workbook.add_worksheet("Predicter")

try:
    csv_file = open(latestFile)
    csv_reader = csv.DictReader(csv_file)

    teams = {}

    # data cell format
    cellfmt = workbook.add_format()
    cellfmt.set_text_wrap()
    cellfmt.set_align('center')
    #cellfmt.   ## cell height and width

    # header cell format
    headerfmt = workbook.add_format()
    headerfmt.set_text_wrap()
    headerfmt.set_align('center')
    headerfmt.set_bold()
    headerfmt.set_bg_color('gray')

    # Title cell format
    titlefmt = workbook.add_format()
    headerfmt.set_text_wrap()
    headerfmt.set_bg_color('gray')
    headerfmt.set_font_size(30)

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
            r=0
            teams[row["Team Number"]].write(r, 0, teamNum, cellfmt)
            r+=1
            headerRow=r
            teams[row["Team Number"]].write(headerRow, 0, "Event", cellfmt)

        col=1
        r+=1

        # Loop through each column in the next row
        for c, (k, v) in enumerate(sorted(row.items(), key=getOrder)):
            # if the key starts with a space, ignore it, it is a header
            if re.match('^ ',k): 
                continue
            # skip the team number column, we are on the team's sheet
            if re.match('Team Number',k): 
                continue
            # remove the quotes (should this be part of the csv reader?)
            if type(k) is str: k = re.sub('["]', '', k)
            if type(v) is str: v = re.sub('["]', '', v)
            # Add header in top row
            teams[row["Team Number"]].write(headerRow, col, k, cellfmt)
            # Add in the event name in the first column
            teams[row["Team Number"]].write(r, 0, event, cellfmt)
            # Add values to rows
            teams[row["Team Number"]].write(r, col, v, cellfmt)
            col+=1

except Exception as ee:
    print(ee)

finally:
    csv_file.close()
    
# add all the raw datat to a sheet.
with open (latestFile) as csv_file:
    rawData = workbook.add_worksheet("Raw Data")
    reader = csv.reader(csv_file)
    for r, rows in enumerate(reader):
        for c, col in enumerate(rows):
            rawData.write(r, c, col)

workbook.close()
