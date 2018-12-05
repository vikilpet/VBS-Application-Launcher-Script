' ������ ��������� � �������� �� ������ �������������
' 2018-06-11
' 2018-07-11 + EventToLog
' 2018-07-12 + ������ ����������� + ����� ���������� + ���������
' 2018-07-13_+ Duplicate + Const
' 2018-12-02 + BLOCK

Option Explicit

Dim K, strUser, blnFound, objShell, APP_USERLIST

' ++++++++++ SETTINGS ++++++++++
	Const APP_NAME = "1� Accounting"
	APP_USERLIST = Array("admin", "user1", "user2")
	Const APP_PATH = """c:\Program Files\1cv8\8.3.13.1644\bin\1cv8c.exe"" ENTERPRISE /F""f:\1c_acc"""
	Const APP_EVENT = "Account"
	Const BLOCK = ""
	Const LOG_FILE = "e:\apps.log"
	Const DUP_STRING = "d:\1c_acc"
	Const DRIVE_LETTER = "Z:"
	' ��� ��������� ����� �� �����:
	Const DRIVE_PATH = "\\192.168.0.1\share"
' ---------- SETTINGS ----------

On Error Resume Next

If Len(BLOCK) > 0 Then
	MsgBox "������ ���������� �" & APP_NAME & "� ������������ �� �������:" & _
	vbCrLf & "�" & BLOCK & "�", _
		vbExclamation, "������"
	WScript.Quit
End If

If Duplicate(DUP_STRING) Then
	EventToLog LOG_FILE, APP_EVENT & vbTab & "duplicate"
	MsgBox "���������� �" & APP_NAME & "� � ��� ��� ��������", _
		vbExclamation, "������"
	WScript.Quit
End If

strUser = CreateObject("WScript.Network").UserName

blnFound = False

For K = lbound(APP_USERLIST) to ubound(APP_USERLIST)
	If strUser = APP_USERLIST(K) Then
		blnFound = True
		Exit For
	End If
Next

If blnFound Then
	' MsgBox APP_NAME

	EventToLog LOG_FILE, APP_EVENT & vbTab & "launch"

	Set objShell = CreateObject( "WScript.Shell" )
	objShell.Run APP_PATH
	If Err.Number = 0 Then
		MapZDrive DRIVE_LETTER, DRIVE_PATH
	Else
		EventToLog LOG_FILE, APP_EVENT & vbTab & "failed to launch"
		
		MsgBox "������ ������� ���������� " & APP_NAME & _
			" ��� ������������ �" & strUser & "�:" & vbCrLf & vbCrLf & _
			Err.Description & " (" & Err.Number & ")", vbCritical, "������� ����!"
	End If
Else
	EventToLog LOG_FILE, APP_EVENT & vbTab & "denied"
	MsgBox "� ��� (�" & strUser & _
		"�) ��� ���� �� ������ ��������� �" & APP_NAME & _
		"�", vbExclamation, "������"
End If


Function EventToLog(strFile, APP_EVENT)
	' 2018-07-12_17-17-58
	' 2018-07-25_17-58-45 + ClientName

	Dim objNetwork, t, strTimeStamp, objFSO, objLog, strNetUser, _
	strComputer, strDomain, strClientName
	
	strNetUser = CreateObject("WScript.Network").UserName
	' strNetUser = objNetwork.UserName
	' �� ������������ ������� � ��� ��� ������ �������, ����������:
	' strComputer = objNetwork.ComputerName
	' strDomain = objNetwork.UserDomain
	
	strClientName = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%ClientName%")

	t = Now()
	strTimeStamp = Year(t) & "-" & _
	Right("0" & Month(t),2)  & "-" & _
	Right("0" & Day(t),2)  & "_" & _  
	Right("0" & Hour(t),2) & "-" & _
	Right("0" & Minute(t),2) & "-" & _
	Right("0" & Second(t),2) 

	Set objFSO = CreateObject("Scripting.FileSystemObject")

	' ���������: fullpath, 8 - ������� ��� ��������, True - ������� ����, ���� �� ����������
	Set objLog = objFSO.OpenTextFile(strFile, 8, True)
	
	objLog.WriteLine strTimeStamp & vbTab & _
	strNetUser & vbTab & strClientName & vbTab & APP_EVENT

	objLog.Close

End Function


Function MapZDrive(strDriveLetter, strDrivePath)
	Dim objFSO, objNetwork
	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	If Not objFSO.DriveExists(strDriveLetter) Then
		Set objNetwork = CreateObject("WScript.Network")
		objNetwork.MapNetworkDrive strDriveLetter, strDrivePath
	Else
		'MsgBox "Disk is already mapped"
	End If
	  
End Function


Function Duplicate(strSearch)
	Dim objWMIService, strQuery, ColItems, objItem, blnResult

	Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
	strQuery = "Select * from Win32_Process"
	set ColItems = objWMIService.ExecQuery(strQuery,,48)
	
	blnResult = False
	
	For Each objItem in colItems
		' � ������ (���������?) ��������� ��� Null, 
		If InStr(objItem.CommandLine, strSearch) Then
			'MsgBox "FOUND " & objItem.Name & " " & objItem.CommandLine
			blnResult = True
		End If
	Next
	
	Duplicate = blnResult
	
End Function











