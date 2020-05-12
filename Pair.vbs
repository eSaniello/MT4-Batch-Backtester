data_dir_path = "C:\Users\Shaniel Samadhan\AppData\Roaming\MetaQuotes\Terminal\1DAFD9A7C67DC84FE37EAA1FC1E5CF75\"
base_dir_path = "C:\Program Files (x86)\MetaTrader 4 IC Markets\"
terminal_path = base_dir_path & "terminal.exe"
output_dir_path = "reports\" 'cannot be an absolute path, it is a relative to data_dir_path

file_name = "01" '.ini'
expert = "ALGO 1 backtest" 'ea name'

settings_file_tmplt = data_dir_path & file_name & ".set"
ini_file_tmplt = data_dir_path & file_name & ".ini"

'("AUDCAD","AUDCHF","AUDJPY","AUDNZD","AUDUSD","CADCHF","CADJPY","CHFJPY","EURCHF","EURAUD","EURCAD","EURGBP","EURJPY","EURNZD","EURUSD")
'("GBPAUD","GBPCAD","GBPCHF","GBPJPY","GBPNZD","GBPUSD","NZDCHF","NZDCAD","NZDJPY","NZDUSD","USDCAD","USDCHF","USDJPY","XAUUSD","XAGUSD")

symbol_arr = Array("EURUSD", "AUDNZD", "EURGBP", "AUDCAD", "CHFJPY")

start_date = "2017.01.01"
end_date = "2020.05.12"

build_report = 1
build_data= 1

' BEGIN - DO NOT EDIT NEXT LINES
symbol_token = "<SYMBOL>"
start_date_token = "<START_DATE>"
end_date_token = "<END_DATE>"
file_token = "<FILE>"
report_token = "<REPORT>"
expert_token = "<EXPERT>"

Set fso = CreateObject("Scripting.FileSystemObject")

For Each symbol In symbol_arr
	current_report_dir = data_dir_path & output_dir_path & expert & "\" & symbol
	data_path = data_dir_path & output_dir_path & expert & "\" & symbol & "\data " & Mid(start_date, 1, 4) & "-" & Mid(end_date, 1, 4) & ".tab"
	
	If fso.FolderExists(current_report_dir) And build_report = 1 Then
		Wscript.Echo "Found folder: [" & current_report_dir & "]"
		Wscript.Echo "Keep (K) or delete (D)?"
		answer = Wscript.StdIn.ReadLine
		If UCase(answer) = "D" Then
			fso.DeleteFolder current_report_dir
			Wscript.Echo "Folder deleted: [" & current_report_dir & "]"
		End If
	End If 
	
	BuildFullPath current_report_dir
	
	If fso.FileExists(data_path) And build_data = 1 Then
		Wscript.Echo "Found data file: [" & data_path & "]"
		Wscript.Echo "Keep (K) or Delete (D)?"
		answer = Wscript.StdIn.ReadLine
		If UCase(answer) = "D" Then
			fso.DeleteFile data_path
			Wscript.Echo "File deleted: [" & data_path & "]"
		End If
	End If
	
	set_file = "VPU Algo " & symbol & " " & Mid(start_date, 1, 4) & "-" & Mid(end_date, 1, 4)
	ini_file = set_file
	report_file = output_dir_path & expert & "\" & symbol & "\" & set_file
	
	fso.CopyFile settings_file_tmplt, data_dir_path & "tester\" & set_file & ".set"
	fso.CopyFile ini_file_tmplt, data_dir_path & "tester\" & ini_file & ".ini"
	
	ReplaceInFile data_dir_path & "tester\" & ini_file & ".ini", symbol_token, symbol
	ReplaceInFile data_dir_path & "tester\" & ini_file & ".ini", start_date_token, start_date
	ReplaceInFile data_dir_path & "tester\" & ini_file & ".ini", end_date_token, end_date
	ReplaceInFile data_dir_path & "tester\" & ini_file & ".ini", file_token, set_file
	ReplaceInFile data_dir_path & "tester\" & ini_file & ".ini", report_token, report_file
	ReplaceInFile data_dir_path & "tester\" & ini_file & ".ini", expert_token, expert
	
	If build_report = 1 Then
		Set oShell = WScript.CreateObject("WScript.Shell")
		command = "cmd /C " & Chr(34) & Chr(34) & terminal_path & Chr(34) & " " & Chr(34) & data_dir_path & "tester\" & ini_file & ".ini" & Chr(34) & " /skipupdate" & Chr(34)
		Wscript.Echo command
		oShell.Run command, 0, true
	End If
	
	If build_data = 1 Then
		GenerateData data_dir_path & report_file & ".htm", data_path, set_file
	End If

	fso.DeleteFile data_dir_path & "tester\" & set_file & ".set"
	fso.DeleteFile data_dir_path & "tester\" & ini_file & ".ini"
Next

Sub GenerateData(ReportFilePath, OutFilePath, Description)
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set doc = CreateObject("HTMLFILE")
	Set file = fso.OpenTextFile(ReportFilePath, 1)
	
	file_empty = 0
	If Not fso.FileExists(OutFilePath) Then
		Set outfile = fso.CreateTextFile(OutFilePath, True)
		file_empty = 1
	Else
		Set outfile = fso.OpenTextFile(OutFilePath, 8)
	End If

	xmlString = file.ReadAll	
	xmlString = replace(xmlString, "Relative drawdown</td><td align=right>", "Relative drawdown</td><td id='rdd' align=right>", 1, -1, 0)
	xmlString = replace(xmlString, "Total net profit</td><td align=right>", "Total net profit</td><td id='tnp' align=right>", 1, -1, 0)
	xmlString = replace(xmlString, "Total trades</td><td align=right>", "Total trades</td><td id='tt' align=right>", 1, -1, 0)
	xmlString = replace(xmlString, "Profit trades (% of total)</td><td align=right>", "Profit trades (% of total)</td><td id='pt' align=right>", 1, -1, 0)
	
	doc.write xmlString
	
	Set node = doc.getElementById("rdd")
	rdd = node.InnerHTML
	rdd = Mid(rdd, 1, InStr(rdd, "%") - 1)

	Set node = doc.getElementById("tnp")
	tnp = node.InnerHTML

	Set node = doc.getElementById("tt")
	tt = node.InnerHTML
	
	Set node = doc.getElementById("pt")
	pt = node.InnerHTML
	pt = Mid(pt, InStr(pt, "(") + 1, InStr(pt, ")") - InStr(pt, "(") - 2)

	If file_empty = 1 Then
		outfile.WriteLine "Description" & Chr(9) & "Relative DD" & Chr(9) & "Net Profit" & Chr(9) & "Trades" & Chr(9) & "Win Rate"
	End If
	
	outfile.WriteLine Description & Chr(9) & rdd & Chr(9) & tnp & Chr(9) & tt & Chr(9) & pt
    outfile.Close
End Sub	

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
	Set fso = CreateObject("Scripting.FileSystemObject")
	
	If Not fso.FolderExists(FullPath) Then
		BuildFullPath fso.GetParentFolderName(FullPath)
		fso.CreateFolder FullPath
	End If
End Sub

