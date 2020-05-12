data_dir_path = "C:\Users\Shaniel Samadhan\AppData\Roaming\MetaQuotes\Terminal\1DAFD9A7C67DC84FE37EAA1FC1E5CF75\"
base_dir_path = "C:\Program Files (x86)\MetaTrader 4 IC Markets\"
terminal_path = base_dir_path & "terminal.exe"
output_dir_path = "reports\" 'cannot be an absolute path, it is a relative to data_dir_path

'file_name = "Algo"
'expert = "01"

file_name = "01" '.ini'
expert = "ALGO 1 backtest" 'ea name'

settings_file_tmplt = data_dir_path & file_name & ".set"
ini_file_tmplt = data_dir_path & file_name & ".ini"

symbol_arr = Array("EURUSD")

start_date = "2017.01.01"
end_date = "2020.05.01"
period_start = 5
period_end = 40
period_step = 5



' BEGIN - DO NOT EDIT NEXT LINES
symbol_token = "<SYMBOL>"
period_token = "<PERIOD>"
start_date_token = "<START_DATE>"
end_date_token = "<END_DATE>"
file_token = "<FILE>"
report_token = "<REPORT>"
expert_token = "<EXPERT>"

Set fso = CreateObject("Scripting.FileSystemObject")

For Each symbol In symbol_arr
	current_report_dir = data_dir_path & output_dir_path & expert & "\" & symbol
	
	If fso.FolderExists(current_report_dir) Then 
		fso.DeleteFolder current_report_dir
		Wscript.Echo "Folder deleted: [" & current_report_dir & "]"
	End If 
	
	BuildFullPath current_report_dir
	
	For period = period_start To period_end Step period_step
		set_file = file_name & " " & period & " " & symbol & " " & Mid(start_date, 1, 4) & "-" & Mid(end_date, 1, 4)
		ini_file = set_file
		report_file = output_dir_path & expert & "\" & symbol & "\" & set_file
		
		fso.CopyFile settings_file_tmplt, data_dir_path & "tester\" & set_file & ".set"
		fso.CopyFile ini_file_tmplt, data_dir_path & "tester\" & ini_file & ".ini"
		
		ReplaceInFile data_dir_path & "tester\" & set_file & ".set", period_token, period
		
		ReplaceInFile data_dir_path & "tester\" & ini_file & ".ini", symbol_token, symbol
		ReplaceInFile data_dir_path & "tester\" & ini_file & ".ini", start_date_token, start_date
		ReplaceInFile data_dir_path & "tester\" & ini_file & ".ini", end_date_token, end_date
		ReplaceInFile data_dir_path & "tester\" & ini_file & ".ini", file_token, set_file
		ReplaceInFile data_dir_path & "tester\" & ini_file & ".ini", report_token, report_file
		ReplaceInFile data_dir_path & "tester\" & ini_file & ".ini", expert_token, expert
		
		Set oShell = WScript.CreateObject("WScript.Shell")
		command = "cmd /C " & Chr(34) & Chr(34) & terminal_path & Chr(34) & " " & Chr(34) & data_dir_path & "tester\" & ini_file & ".ini" & Chr(34) & " /skipupdate" & Chr(34)
		Wscript.Echo command
		oShell.Run command, 0, true
	Next
Next

Sub ReplaceInFile(StrFilename, StrSearch, StrReplace)
	'Does file exist?
	Set fso = CreateObject("Scripting.FileSystemObject")

	'Read file
	set objFile = fso.OpenTextFile(StrFilename, 1)
	oldContent = objFile.ReadAll
	 
	'Write file
	newContent = replace(oldContent, StrSearch, StrReplace, 1, -1, 0)
	set objFile = fso.OpenTextFile(StrFilename, 2)
	objFile.Write newContent
	objFile.Close
End Sub

Sub BuildFullPath(ByVal FullPath)
	If Not fso.FolderExists(FullPath) Then
		BuildFullPath fso.GetParentFolderName(FullPath)
		fso.CreateFolder FullPath
	End If
End Sub

