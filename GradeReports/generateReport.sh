#!/bin/bash
#
# We have to stop LibreOffice since we need to open it with some additional
# command line arguments to make it listen on a port that we can connect with
# in Python. Then we'll run the Python script to update the report of
# everybody's grades.
#
# Make sure you set "weigthed" and "gradeSheets" in the Python file.
#
if ps -ef | grep -v grep | grep soffice &>/dev/null; then
    echo "Please close LibreOffice first."
    exit 1
fi

repo="/home/garrett/Documents/Github/HighSchoolMath"
file="GradesQ1_points.ods"

soffice --calc --accept="socket,host=localhost,port=2002;urp;StarOffice.ServiceManager" \
    "$repo/$file" &

echo "Press enter to continue."
read -r line

python3 "$repo"/GradeReports/GenerateReport.py
