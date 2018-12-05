' Запуск программы с доступом по списку пользователей
' 2018-06-11
' 2018-07-11 + EventToLog
' 2018-07-12 + больше логирования + имена переменных + сообщения
' 2018-07-13_+ Duplicate + Const
' 2018-12-02 + BLOCK

Option Explicit

Dim K, strUser, blnFound, objShell, APP_USERLIST

' ++++++++++ SETTINGS ++++++++++
	Const APP_NAME = "1с Accounting"
	APP_USERLIST = Array("admin", "user1", "user2")
	Const APP_PATH = """c:\Program Files\1cv8\8.3.13.1644\bin\1cv8c.exe"" ENTERPRISE /F""f:\1c_acc"""
	Const APP_EVENT = "Account"
	Const BLOCK = ""
	Const LOG_FILE = "e:\apps.log"
	Const DUP_STRING = "d:\1c_acc"
	Const DRIVE_LETTER = "Z:"
	' Без обратного слеша на конце:
	Const DRIVE_PATH = "\\192.168.0.1\share"
' ---------- SETTINGS ----------

On Error Resume Next

If Len(BLOCK) > 0 Then
	MsgBox "Запуск приложения «" & APP_NAME & "» заблокирован по причине:" & _
	vbCrLf & "«" & BLOCK & "»", _
		vbExclamation, "Сервер"
	WScript.Quit
End If

If Duplicate(DUP_STRING) Then
	EventToLog LOG_FILE, APP_EVENT & vbTab & "duplicate"
	MsgBox "Приложение «" & APP_NAME & "» у вас уже запущено", _
		vbExclamation, "Сервер"
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
		
		MsgBox "Ошибка запуска приложения " & APP_NAME & _
			" для пользователя «" & strUser & "»:" & vbCrLf & vbCrLf & _
			Err.Description & " (" & Err.Number & ")", vbCritical, "Позвони Илье!"
	End If
Else
	EventToLog LOG_FILE, APP_EVENT & vbTab & "denied"
	MsgBox "У вас («" & strUser & _
		"») нет прав на запуск программы «" & APP_NAME & _
		"»", vbExclamation, "Сервер"
End If


Function EventToLog(strFile, APP_EVENT)
	' 2018-07-12_17-17-58
	' 2018-07-25_17-58-45 + ClientName

	Dim objNetwork, t, strTimeStamp, objFSO, objLog, strNetUser, _
	strComputer, strDomain, strClientName
	
	strNetUser = CreateObject("WScript.Network").UserName
	' strNetUser = objNetwork.UserName
	' На терминальном сервере в них имя самого сервера, бесполезны:
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

	' Параметры: fullpath, 8 - открыть для дозаписи, True - создать файл, если не существует
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
		' У многих (системных?) процессов она Null, 
		If InStr(objItem.CommandLine, strSearch) Then
			'MsgBox "FOUND " & objItem.Name & " " & objItem.CommandLine
			blnResult = True
		End If
	Next
	
	Duplicate = blnResult
	
End Function











