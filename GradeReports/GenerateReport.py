"""
Generate grade reports

It'll generate reports in this directory (so make sure you're in the proper
directory when you run this script, or use the generateReport.sh script).
Be sure to change "weighted" and "gradeSheets" variables based on whether you
pick the points vs. weighted template and what the classes are that you teach
(the sheet names in the LibreOffice Calc file).

Run LibreOffice with:
    soffice --calc --accept="socket,host=localhost,port=2002;urp;StarOffice.ServiceManager" /path/to/GradesQ1.ods

(see the generateReport.sh script)
"""
import socket  # only needed on win32-OOo3.0.0
import uno
import string # for string.ascii_uppercase
import os
import datetime
from operator import itemgetter

#
# Options
#
# CHANGE THESE!

# Do we have a points-based or weighted grading document?
weighted = False
# Names of the sheets in the file
gradeSheets = ["11th", "8th", "9th", "12th,PC", "12th,Comp", "10th"]

#
# Connect to LibreOffice
#

# get the uno component context from the PyUNO runtime
localContext = uno.getComponentContext()

# create the UnoUrlResolver
resolver = localContext.ServiceManager.createInstanceWithContext(
                        "com.sun.star.bridge.UnoUrlResolver", localContext )

# connect to the running office
ctx = resolver.resolve( "uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext" )
smgr = ctx.ServiceManager

# get the central desktop object
desktop = smgr.createInstanceWithContext( "com.sun.star.frame.Desktop",ctx)

# access the current writer document
model = desktop.getCurrentComponent()

#
# Setup for generating the report
#

# Get the different sheets we'll be working with
grades = [model.Sheets.getByName(i) for i in gradeSheets]
report = model.Sheets.getByName("report")

# Find the start and end rows of the different sections
def findSection(sheet, start, end="-", column=1):
    row = 0
    startRow = None
    endRow = None
    startColumn = 1

    # If start==None, then points-based and it starts at the second row
    # (i.e. row 1). However, we start after the startRow ignoring it since
    # normally it contains the text like "Test/Quizzes," so here we'll start
    # *after* row 0.
    if start == None:
        startRow = 0

    while startRow == None or endRow == None:
        cell = sheet.getCellByPosition(column, row).getString()

        if startRow == None:
            if cell == start:
                startRow = row
        else:
            if cell == end:
                endRow = row

        row += 1

    return [startRow+1,endRow-1]

# What is in a section? e.g. the list of all tests/quizzes
# [[row 3, "HW 1.3", out of 15.2 points, "HW 1.3: 1, 2, 3, ...", date], ...]
def sectionContents(sheet, rowRange, column=1, outOfColumn=2, dateColumn=0):
    contents = []

    for row in range(rowRange[0], rowRange[1]+1):
        cell = sheet.getCellByPosition(column, row)
        outOf = str(round(sheet.getCellByPosition(outOfColumn, row).getValue(),1))
        dateStr = sheet.getCellByPosition(dateColumn, row).getString()
        date = datetime.datetime.strptime(dateStr, "%m/%d/%Y").date()
        s_full = cell.getString()

        # Strip all after the colon
        s_part = s_full.split(":")[0]
        s_full = s_full.split("[")[0]

        contents.append([row,s_part,outOf,s_full,date])

    return contents

# Get the list of students
#
# Returns [[6, "Billy"], [7, "Susan"], ...] where 6, 7, etc. are column
# numbers
def getStudents(sheet, row="2"):
    students = []

    # Student names in F2 through AZ2
    for column in list(string.ascii_uppercase[5:])+["A"+i for i in list(string.ascii_uppercase)[0:11]]:
        name = sheet.getCellRangeByName(column+row)
        studentName = name.getString()

        if studentName and studentName[0:12] != "Student Name":
            students.append([name.getCellAddress().Column, studentName])

    return students

def getTestGrade(sheet, column, rowRange):
    return sheet.getCellByPosition(column, rowRange[1]+4).getValue()

def getAssignmentsGrade(sheet, column, rowRange):
    return sheet.getCellByPosition(column, rowRange[1]+5).getValue()

def getParticipationGrade(sheet, column, rowRange):
    return sheet.getCellByPosition(column, rowRange[1]+4).getValue()

def getTotalGrade(sheet, column, rowRange):
    # TODO remove this check
    #if sheet == grades[1]: # 8th hasn't had tests yet
    #    return sheet.getCellByPosition(column, rowRange[1]+7).getValue()
    #else:
    return sheet.getCellByPosition(column, rowRange[1]+6).getValue()

def getLetterGrade(sheet, column, rowRange):
    # TODO remove this check
    #if sheet == grades[1]: # 8th hasn't had tests yet
    #    return sheet.getCellByPosition(column, rowRange[1]+8).getString()
    #else:
    return sheet.getCellByPosition(column, rowRange[1]+7).getString()

def getTotalGradePoints(sheet, column, rowRange):
    return sheet.getCellByPosition(column, rowRange[1]+3).getValue()

def getLetterGradePoints(sheet, column, rowRange):
    return sheet.getCellByPosition(column, rowRange[1]+4).getString()

def getClassName(sheet, column="1", row="0"):
    return sheet.getCellByPosition(column, row).getString()

# 0.945 -> 94.5%
def formatGrade(grade, total=False):
    if total:
        return str(round(grade*100)) + "%"
    else:
        return str(round(grade*100, 1)) + "%"

# Managing position in the report
if weighted:
    reportWidth = 3 # columns
else:
    reportWidth = 1 # columns
position = [0,0] # starting position [column, row]
currentUsedRows = [] # List of how many rows are used for each column

def setReport(position, offset, value):
    global report
    cell = report.getCellByPosition(position[0]+offset[0], position[1]+offset[1])

    if isinstance(value, str):
        cell.String = value
    else:
        cell.Value = value

# Create directory for the individual reports
today = datetime.datetime.today()
todayStr = today.strftime("%Y%m%d")

if not os.path.exists(todayStr):
    os.makedirs(todayStr)

# For each grade
for grade in gradeSheets:
    directory = os.path.join(todayStr, grade)
    if not os.path.exists(directory):
        os.makedirs(directory)

# Clear the old report
r = report.getCellRangeByName("A1:AMJ1048576")
r.clearContents(4)

# Make the report look nice, having two columns
rowformat = "{:<25} {:<25}"
rowspace = "   "
rowheadformat = "{:<29} {:<25}"

if __name__ == "__main__":
    #
    # Note: weighted version not tested for a long time, so it might not work
    # that well
    #
    if weighted:
        # Generate the report
        for gradeNum, grade in enumerate(grades):
            # Used in file path for the individual reports
            gradeName = gradeSheets[gradeNum]

            # Start each new grade at the farthest left
            position[0] = 0

            # Get the list of students
            students = getStudents(grade)

            # Find where the Tests/Quizzes are
            assessment = findSection(grade, "Tests/Quizzes")
            assessmentContents = sectionContents(grade, assessment)
            assignments = findSection(grade, "Assignments")
            assignmentsContents = sectionContents(grade, assignments)
            participation = findSection(grade, "Participation")

            # Create the report for each student
            for (num, (column, name)) in enumerate(students):
                # Get first name in the format "Last, First"
                firstName = name[name.find(",")+2:]
                lastInitial = name[0]

                with open(os.path.join(todayStr, gradeName, firstName+" "+lastInitial+".txt"), "w") as f:
                    # Get section grades
                    testsGrade = formatGrade(getTestGrade(grade, column, assessment))
                    assignmentsGrade = formatGrade(getAssignmentsGrade(grade, column, assignments))
                    participationGrade = formatGrade(getParticipationGrade(grade, column, participation))
                    totalGrade = formatGrade(getTotalGrade(grade, column, participation), total=True)
                    letterGrade = getLetterGrade(grade, column, participation)

                    # Start the offset at zero
                    usedRows = 0

                    print(name, file=f)
                    setReport(position, (0,usedRows), name)
                    usedRows += 1

                    # Tests/Quizzes
                    print(file=f)
                    row = ["Tests/Quizzes", testsGrade]
                    print(rowheadformat.format(*row), file=f)

                    setReport(position, (0,usedRows), "Tests/Quizzes")
                    # TODO remove this check
                    #if grade == grades[1]: # 8th hasn't had tests yet
                    #    setReport(position, (1,usedRows), "N/A")
                    #else:
                    setReport(position, (1,usedRows), testsGrade)
                    usedRows += 1

                    # List of missing work
                    missing = []

                    # Get grades for each of the assessments
                    for (row, contentName, outOf, fullName, date) in assessmentContents:
                        cell = grade.getCellByPosition(column, row)
                        contentGrade = str(round(cell.getValue(),1)) + "/" + outOf
                        row = [contentName, contentGrade]
                        print(rowspace, rowformat.format(*row), file=f)

                        # "Missed work" if a zero and due before today, and not already seen
                        if cell.getValue() == 0 and date < today.date() \
                            and cell.getString() != "late" \
                            and cell.getString() != "cheat":
                            missing.append([contentName, contentGrade, fullName])

                    # Assignments
                    row = ["Assignments", assignmentsGrade]
                    print(rowheadformat.format(*row), file=f)

                    setReport(position, (0,usedRows), "Assignments")
                    setReport(position, (1,usedRows), assignmentsGrade)
                    usedRows += 1

                    # Get grades for each of the assignments
                    for (row, contentName, outOf, fullName, date) in assignmentsContents:
                        cell = grade.getCellByPosition(column, row)
                        contentGrade = str(round(cell.getValue(),1)) + "/" + outOf
                        row = [contentName, contentGrade]
                        print(rowspace, rowformat.format(*row), file=f)

                        # "Missed work" if a zero and due before today, and not already seen
                        if cell.getValue() == 0 and date < today.date() \
                            and cell.getString() != "late" \
                            and cell.getString() != "cheat":
                            missing.append([contentName, contentGrade, fullName])

                    # Participation (summary)
                    row = ["Participation", participationGrade]
                    print(rowheadformat.format(*row), file=f)

                    setReport(position, (0,usedRows), "Participation")
                    setReport(position, (1,usedRows), participationGrade)
                    usedRows += 1

                    # Total grade
                    print(file=f)
                    print("Total:", totalGrade, "=", letterGrade, file=f)
                    setReport(position, (0,usedRows), "Total")
                    setReport(position, (1,usedRows), totalGrade + " = " + letterGrade)
                    usedRows += 1

                    # Missing work
                    print(file=f)
                    print("Missing Work", file=f)
                    missingWork = 0
                    for (contentName, contentGrade, fullName) in missing:
                        # Corrections and Extra Credit aren't required
                        if contentName != "Corrections" and contentName[-2:] != "EC":
                            print(rowspace, fullName, file=f)
                            missingWork += 1

                    if missingWork == 0:
                        print(rowspace, "None", file=f)

                    # Manage which column/row we are going to put this in
                    position[0] = ((num+1)%reportWidth)*2 # 2 columns/student
                    currentUsedRows.append(usedRows)

                    # If we're starting a new row
                    if position[0] == 0 or num+1 == len(students):
                        position[1] += max(currentUsedRows)+1
                        currentUsedRows.clear()

    else:
        # Generate the report
        for gradeNum, grade in enumerate(grades):
            # Used in file path for the individual reports
            gradeName = gradeSheets[gradeNum]

            # Start each new grade at the farthest left
            position[0] = 0

            # Put a description of what class this is first
            setReport(position, (0,0), getClassName(grade))
            position[1] += 2 # Two lines to make it spaced out a bit
            #setReport(position, (0,0), "Student")
            #setReport(position, (1,0), "Grade")
            #position[1] += 1

            # Get the list of students
            students = getStudents(grade, "1")

            # Get points-based grading
            points = findSection(grade, None)
            pointsContents = sectionContents(grade, points)

            # List of student name and grades to put into the Report in Calc
            data = []

            # Create the report for each student
            for (num, (column, name)) in enumerate(students):
                # Get first name in the format "Last, First"
                firstName = name[name.find(",")+2:]
                lastInitial = name[0]

                with open(os.path.join(todayStr, gradeName, firstName+" "+lastInitial+".txt"), "w") as f:
                    # Get section grades
                    totalGrade = formatGrade(getTotalGradePoints(grade, column, points), total=True)
                    letterGrade = getLetterGradePoints(grade, column, points)

                    # Student names
                    print(name, file=f)
                    print(file=f)

                    # List of missing work
                    missing = []

                    # Get grades for each of the graded items
                    for (row, contentName, outOf, fullName, date) in pointsContents:
                        cell = grade.getCellByPosition(column, row)
                        contentGrade = str(round(cell.getValue(),1)) + "/" + outOf
                        row = [contentName, contentGrade]
                        print(rowspace, rowformat.format(*row), file=f)

                        # "Missed work" if a zero and due before today, and not already seen
                        #  - If it's late, then the missed work was still turned in
                        #  - Not really sure what should be done with cheating...
                        #  - Bellwork and Instant Quizzes can't be made up
                        #  - Corrections and Extra Credit aren't required
                        if cell.getValue() == 0 and date < today.date() \
                            and cell.getString() != "late" \
                            and cell.getString() != "cheat" \
                            and contentName != "Bellwork" \
                            and contentName != "Instant Quiz" \
                            and contentName != "Corrections" \
                            and contentName[-2:] != "EC":
                            missing.append([contentName, contentGrade, fullName])

                    # Total grade
                    print(file=f)
                    print("Total:", totalGrade, "=", letterGrade, file=f)

                    # Missing work
                    print(file=f)
                    print("Missing Work", file=f)
                    missingWork = 0
                    for (contentName, contentGrade, fullName) in missing:
                        print(rowspace, fullName, file=f)
                        missingWork += 1

                    if missingWork == 0:
                        print(rowspace, "None", file=f)

                    data.append((name, totalGrade + " = " + letterGrade))

            # Sort data by student last name
            data = sorted(data, key=itemgetter(0))

            # Data for Report
            for (name, gradeStr) in data:
                # Start the offset at zero
                usedRows = 0

                setReport(position, (0,usedRows), name)
                setReport(position, (1,usedRows), gradeStr)

                # Manage which column/row we are going to put this in
                position[0] = ((num+1)%reportWidth)*2 # 2 columns/student
                currentUsedRows.append(usedRows)

                # If we're starting a new row
                if position[0] == 0 or num+1 == len(students):
                    position[1] += max(currentUsedRows)+1
                    currentUsedRows.clear()

            # Add a little space after each class
            position[1] += 1
