# $language = "VBScript"
# $interface = "1.0"

'|======================================================================================| 
'|                                      HOWTO                                           |
'|======================================================================================|
'|1. Run the script at SecureCRT ---> Script ---> Run --> Choose Script                 |
'|2. Wait for a while, when prompted choose Command File                    			|
'|3. Wait for a while, when prompted enter your Username & Password                    	|
'|======================================================================================|

Sub Main
	sIniDir = ""
	sFilter = "Text File (*.txt)|*.txt|"
	sTitle = "Choose Command File"
	
	'Show File Dialog And Get Textfile
	readtxt = GetFileDlg(Replace(sIniDir,"\","\\"),sFilter,sTitle)
  
	If readtxt = "" Then
		Exit Sub
	End If
	
	crt.Screen.Synchronous = True
	
	'Run Command From Textfile
	RunCommand(readtxt)
	
	crt.Screen.Synchronous = False
End Sub

Function GetFileDlg(sIniDir,sFilter,sTitle)
  GetFileDlg=CreateObject("WScript.Shell").Exec("mshta.exe ""about:<object id=d classid=clsid:3050f4e1-98b5-11cf-bb82-00aa00bdce0b></object><script>moveTo(0,-9999);function window.onload(){var p=/[^\0]*/;new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(1).Write(p.exec(d.object.openfiledlg('" & sIniDir & "',null,'" & sFilter & "','" & sTitle & "')));close();}</script><hta:application showintaskbar=no />""").StdOut.ReadAll
End Function

Sub RunCommand(readtxt)
	Dim fso, file, data, vWaitFors, splitPrompt
	
	Set fso = CreateObject("Scripting.FileSystemObject") 
	Set file = fso.OpenTextFile (readtxt, 1)
	
	row = crt.Screen.CurrentRow
	strDesc = crt.Screen.Get(row, 1, row, 256)
	
	If InStr(strDesc, ":") > 0 Then
		splitPrompt = Split(strDesc, ":")
		realhostname = splitPrompt(1)
		If InStr(strDesc, "#") > 0 Then
			splitPrompt = Split(realhostname, "#")
			realhostname = splitPrompt(0)
		End If
	End If
	
	If InStr(strDesc, "#") > 0 Then
		splitPrompt = Split(strDesc, "#")
		realhostname = splitPrompt(0)
	End If
	
	If InStr(strDesc, "@") > 0 Then
		splitPrompt = Split(strDesc, "@")
		realhostname = splitPrompt(1)
		If InStr(strDesc, ">") > 0 Then
			splitPrompt = Split(realhostname, ">")
			realhostname = splitPrompt(0)
		End If
		If InStr(strDesc, "<") > 0 Then
			splitPrompt = Split(realhostname, "<")
			realhostname = splitPrompt(1)
		End If
	End If
	
	If InStr(strDesc, ">") > 0 Then
		splitPrompt = Split(strDesc, ">")
		realhostname = splitPrompt(0)
		If InStr(strDesc, "<") > 0 Then
			splitPrompt = Split(realhostname, "<")
			realhostname = splitPrompt(1)
		End If
	End If
	
	If InStr(strDesc, "[") > 0 Then
		splitPrompt = Split(strDesc, "[")
		realhostname = splitPrompt(1)
		If InStr(strDesc, "]") > 0 Then
			splitPrompt = Split(realhostname, "]")
			realhostname = splitPrompt(0)
		End If
	End If
	
	vWaitFors = Array(realhostname & "#", realhostname & ">", "[" & realhostname, "--More--")
	
	Do Until file.AtEndOfStream
		command = LTrim(RTrim(Replace(file.Readline, vbTab, " ")))
		If command <> "" Then
			cmd = ""
			For i = 1 to Len(command)
				character = Mid(command, i, 1)
				If i = 1 Then
					cmd = cmd & character
				Else
					prev_character = Mid(command, i-1, 1)
					If (character <> " ") Or (character = " " And prev_character <> " ") Then
						cmd = cmd & character
					End If
				End If
			Next
			If cmd <> "" Then
				row = crt.Screen.CurrentRow
				strDesc = crt.Screen.Get(row, 1, row, 256)
				
				wf1 = realhostname & "#"
				wf2 = "<" & realhostname & ">"
				If InStr(strDesc, "@" & realhostname & "#") > 0 Or InStr(strDesc, "@" & realhostname & ">") > 0 Then
					wf1 = "@" & realhostname & "#"
					wf2 = "@" & realhostname & ">"
				End If
				If InStr(strDesc, realhostname & "#") > 0 Or InStr(strDesc, realhostname & ">") > 0 Then
					wf1 = realhostname & "#"
					wf2 = realhostname & ">"
				End If
				
				strDesc = ""
				row = crt.Screen.CurrentRow
				Do Until strDesc <> ""
					row = row - 1
					strDesc = crt.Screen.Get(row, 1, row, 256)
				Loop

				If InStr(UCase(strDesc), "CONNECT TO HOST") = 0 And InStr(UCase(strDesc), "HOST KEY VERIFICATION FAILED") = 0 Then
					crt.Screen.Send(cmd)
					crt.Screen.WaitForString(cmd)
					crt.Screen.Send(vbCr)
					vWaitFors = Array(wf1, wf2, "[" & realhostname, "[~" & realhostname, "[*" & realhostname, "user#", "console#", "user$", "console$", "bash-3.2$", "profile#", "profile$", "entry#", "entry$", "syslog#", "syslog$", "log-id#", "log-id$")
					crt.Screen.WaitForStrings(vWaitFors)
					
					'Do
					'	nResult = crt.Screen.WaitForStrings(vWaitFors, 2)
					'	If nResult = 0 Then
					'		Exit Do
					'	End If
					'Loop
				End If
				
				row = crt.Screen.CurrentRow
				strDesc = crt.Screen.Get(row, 1, row, 256)
				If InStr(strDesc, "--More--") > 0 Then
					Do Until InStr(strDesc, "--More--") = 0
						crt.Screen.Send " "
						crt.Screen.WaitForStrings(vWaitFors)
						row = crt.Screen.CurrentRow
						strDesc = crt.Screen.Get(row, 1, row, 256)
					Loop
				End If
				
				'ALU - Check Error Command
				row = crt.Screen.CurrentRow - 1
				strDesc = crt.Screen.Get(row, 1, row, 256)
				If InStr(UCase(strDesc), "ERROR: BAD COMMAND") > 0 Then
					MsgBox "Bad Command, check your MOP"
					Exit Do
				End If
				
				If InStr(UCase(strDesc), "MINOR: ") > 0 OR InStr(UCase(strDesc), "ERROR: INVALID PARAMETER") > 0 OR InStr(UCase(strDesc), "ERROR: MISSING PARAMETER") > 0 Then
					MsgBox strDesc
					Exit Do
				End If
				
				'CISCO - Check Error Command
				row = crt.Screen.CurrentRow - 3
				strDesc = crt.Screen.Get(row, 1, row, 256)
				If InStr(UCase(strDesc), "UNKNOWN COMMAND") > 0 Then
					MsgBox "Bad Command, check your MOP"
					Exit Do
				End If
				
				'JUNIPER - Check Error Command
				row = crt.Screen.CurrentRow - 2
				strDesc = crt.Screen.Get(row, 1, row, 256)
				If InStr(UCase(strDesc), "% INVALID INPUT DETECTED AT '^' MARKER") > 0 Then
					MsgBox "Bad Command, check your MOP"
					Exit Do
				End If
				
				'HUAWEI - Check Error Command
				row = crt.Screen.CurrentRow - 1
				strDesc = crt.Screen.Get(row, 1, row, 256)
				If InStr(UCase(strDesc), "ERROR: UNRECOGNIZED COMMAND FOUND AT '^' POSITION") > 0 Then
					MsgBox "Bad Command, check your MOP"
					Exit Do
				End If
				
				strDesc = ""
				row = crt.Screen.CurrentRow
				Do Until strDesc <> ""
					row = row - 1
					strDesc = crt.Screen.Get(row, 1, row, 256)
				Loop
			End If
		End If
	Loop
		
	file.Close
	Set file = Nothing 
End Sub