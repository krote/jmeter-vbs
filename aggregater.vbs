Option Explicit
On Error Resume Next
'*********************************************************************
' Param 1 => 1:Target Multipul Folder, 2:Target One Folder
' Param 2 => Target Folder Path
' Param 3 => Output Folder Path
' ex: Wscript.exe aggregater.vbs 1 c:\jmeter\output c:\jmeter\result
'*********************************************************************

Dim RunnerPath
RunnerPath = "..\lib\ext\CMDRunner.jar"

Dim objParam
Dim inputDir, resultDir, mode

Set objParam = WScript.Arguments

mode = objParam(0)
inputDir = objParam(1)
resultDir = objParam(2)

AddLog "**************************************************"
AddLog "Execute Param:" & mode & "," & inputDir & "," & resultDir

If mode = 1 Then
	ExecRootFolder inputDir, resultDir
Else
	DoOneFolder inputDir, resultDir
End If

Set objParam = Nothing
AddLog "Done!"
MsgBox "Finished"

Err.Clear


'************************************************
'	ExecRootFolder
'************************************************
Sub ExecRootFolder(rootDir, outDir)
	Dim objFSO, folder, objSubFolder

	Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
	If Err.Number = 0 Then
		Set folder = objFSO.GetFolder(rootDir)

		For Each objSubFolder In folder.SubFolders
			AddLog "Searching Folder : " & objSubFolder.Path
			DoOneFolder objSubFolder.Path, outDir
		Next

		Set folder = Nothing
	Else
		AddLog "ERROR:" & Err.Description
	End If
	Set objFSO = Nothing
End Sub

'************************************************
' 	DoOneFolder(
'************************************************
Sub DoOneFolder(inDir, outDir)

	JTL2CSV inDir, inDir

	CollectCSVData inDir, outDir

End Sub


'************************************************
'	JTL2CSV
'************************************************
Sub JTL2CSV(inDir, outDir)
	Dim cmdline, wShell
	Dim objFSO, folder, jtlFiles, jtlFile, result

	Set wShell = CreateObject("WScript.Shell")
	Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
	If Err.Number = 0 Then
		Set folder = objFSO.GetFolder(inDir)
		Set jtlFiles = folder.Files

		For Each jtlFile in jtlFiles
			'WScript.echo jtlFile.Name
			If jtlFile.Type = "JTL ファイル" Then
				AddLog "Exec JTL2CSV => " & jtlFile.Name
				cmdline = "java -jar " & RunnerPath & " --tool Reporter --generate-csv " & outDir & "\result_" & jtlFile.Name & ".csv --input-jtl " & jtlFile.Path & " --plugin-type AggregateReport"
				Set result = wShell.Exec(cmdline)
				Do While result.Status = 0
					WScript.Sleep(100)
				Loop
				AddLog "Exec JTL2CSV => " & jtlFile.Name & " ...  Done"
				Set result = Nothing
			End If
		Next

		Set folder = Nothing
		Set jtlFiles = Nothing
	Else
		AddLog "ERROR:" & Err.Description
	End If

	Set wShell = Nothing
	Set objFSO = Nothing
End Sub

'***************************************************
'	CollectCSVData
'***************************************************
Sub CollectCSVData(inDir, outDir)
	Dim objFSO, folder, csvFiles, csvFile, result
	Dim lines()
	Dim idx, line
	Dim HeaderLine, fFirstLine
	Dim rStream, wStream

	Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")

	If Err.Number = 0 Then
		Set folder = objFSO.GetFolder(inDir)
		Set csvFiles = folder.Files
		idx = 1
		For Each csvFile in csvFiles
			If csvFile.Type = "Microsoft Excel CSV ファイル" Then
				Set rStream = csvFile.OpenAsTextStream()
				fFirstLine = True
				Do Until rStream.AtEndOfStream
					If fFirstLine = True Then
						HeaderLine = "FileName," & rStream.ReadLine
						fFirstLine = False
					Else
						ReDim Preserve lines(idx)
						line = csvFile.Name & "," & rStream.ReadLine
						lines(idx-1) = line
						idx = idx + 1
					End If
				Loop
				rStream.Close
				Set rStream = Nothing
			End If
		Next

		If idx > 1 Then
			Set wStream = objFSO.OpenTextFile(outDir & "\collect_" & folder.Name & ".csv", 2, True)
			wStream.WriteLine HeaderLine
			For idx = 0 To UBound(lines)
				wStream.WriteLine lines(idx)
			Next
			wStream.Close
			Set wStream = Nothing
		End If

		Set folder = Nothing
		Set csvFiles = Nothing
	Else
		AddLog "ERROR:" & Err.Description
	End If

	Set objFSO = Nothing
End Sub

'***************************************************
'	AddLog
'***************************************************
Sub AddLog(strMessage)
	On Error Resume Next
	Const ForAppending = 8 '
	Dim objFSO, logFile, logFileName
	Set objFSO = CreateObject("Scripting.FileSystemObject")

	logFileName = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName) & "\aggregater.log"
	Set logFile = objFSO.OpenTextFile(logFileName, ForAppending, true)
	strMessage = Date() & " " & Time() & ": " & strMessage
	logFile.WriteLine (strMessage)
	Set logFile = Nothing
	Set objFSO = Nothing
End Sub
