Global $userDir = EnvGet("USERPROFILE")
Global $sharedUserDir = EnvGet("PUBLIC")

Func getUserPath($userFileName)
	Return $userDir & "\" & $userFileName
EndFunc
	
Func getUserCloudPath($userFileName)	
	Return $userDir & "\OneDrive\" & $userFileName
EndFunc
	
Func getSharedUserPath($userFileName)	
	Return $sharedUserDir & "\" & $userFileName
EndFunc

Func deleteUserFile($filename)
	If Not FileDelete (getUserPath($filename)) Then
		If Not FileDelete(getUserCloudPath($filename)) Then
			FileDelete(getSharedUserPath($filename))
		EndIf		
	EndIf
EndFunc

Func getUserFilePath($userFileName) 
	Local $userFilePath = getUserPath($userFileName)
	Local $userCloudPath = getUserCloudPath($userFileName)
	Local $sharedUserPath = getSharedUserPath($userFileName)
	if FileExists($userFilePath) Then
		Return $userFilePath
	ElseIf FileExists($userCloudPath) Then
		Return $userCloudPath
	ElseIf FileExists($sharedUserPath) Then
		Return $sharedUserPath
	Else
		Return -1
	EndIf
EndFunc
	
Func diagnoseFile($userFileName)
	Local $userFilePath = getUserFilePath($userFileName)
	if $userFilePath <> -1 Then
		ConsoleWrite($userFilePath & " exists" & @CRLF)
	Else
		ConsoleWrite($userFileName & " not found." & @CRLF)
		
		FileChangeDir($userDir & "\OneDrive\Desktop")
		ConsoleWrite("Searching in " & @WorkingDir & @CRLF)
		Local $hSearch = FileFindFirstFile("http*.*")
		if $hSearch <> -1 Then
			Local $sFileName = "", $iResult = 0
			ConsoleWrite("Found similars" & @CRLF)
			While 1
				$sFileName = FileFindNextFile($hSearch)
				If @error Then ExitLoop
				ConsoleWrite("File: " & $sFileName & @CRLF)
			WEnd
		Else
			ConsoleWrite("nothing similar found" & @CRLF)
		EndIf
	EndIf
	
EndFunc

Func launchProgram($CommandLine)
	ConsoleWrite("Calling '" & $CommandLine & "'" & @CRLF)
	Local $pid = Run($CommandLine, $userDir)
	if not $pid Then
		ConsoleWriteError("Could not launch " & $CommandLine & @CRLF)
		ConsoleWriteError("Error Code is " &  @error & ". Extended " & @extended & @CRLF)
	EndIf
	Return $pid
EndFunc
