' Prints the list of System Restore points to a text file
' © 2005 Ramesh Srinivasan - http://windowsxp.mvps.org
' Last updated on: Aug 20, 2005
' Formats the date / time correctly
'-----------------------------------------------------------

Option Explicit
Dim objWMI, objWMIService, objItem, errResults, clsPoint, strComputer
Dim dtmInstallDate, objOS

' Restore to the point with this description, created with the create_system_restore.bat script
Const RESTORE_POINT = "20160921Lab"

strComputer="."

Set dtmInstallDate = CreateObject( _
"WbemScripting.SWbemDateTime")

Set objWMI = GetObject( _
"winmgmts:\\" & strComputer & "\root\cimv2")

Set objOS = objWMI.ExecQuery( _
"Select * from Win32_OperatingSystem")

Set objWMI = getobject( _
"winmgmts:\\" & strComputer & "\root\default").InstancesOf ("systemrestore")
For Each clsPoint In objWMI
If clsPoint.description = RESTORE_POINT Then
'WScript.Echo "Creation Time= " & getmytime(clsPoint.creationtime)

'WScript.Echo "Description= " & clsPoint.description

'WScript.Echo "Sequence Number= " & clsPoint.sequencenumber

' Restore to this sequence number
Set objWMIService = GetObject("winmgmts:" _
	& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\default")


Set objItem = objWMIService.Get("SystemRestore")
errResults = objItem.Restore(clsPoint.sequencenumber)

End If
Next

function getmytime(wmitime)
dtmInstallDate.Value = wmitime
getmytime = dtmInstallDate.GetVarDate
end function
