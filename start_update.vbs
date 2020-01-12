Dim WShell
Set WShell = WScript.CreateObject("WScript.Shell")

' Windows Update Agent (WUA) base file for offline updates
Dim UpdateBaseURL, UpdateBaseFile
UpdateBaseURL  = "http://go.microsoft.com/fwlink/p/?LinkID=74689"
UpdateBaseFile = "wsusscn2.cab"

' Define the tree of subfolders
Dim CurrentPath, CacheFolder, BaseFolder, UpdatesFolder
CurrentPath   = WShell.CurrentDirectory
CacheFolder   = CurrentPath & "\cache"
BaseFolder    = CacheFolder & "\base"
UpdatesFolder = CacheFolder & "\updates"

Dim BaseFile, DescList
BaseFile = BaseFolder  & "\" & UpdateBaseFile
DescList = CacheFolder & "\updates_desc.txt"

' Define special symbols which can be used with Print() function
Dim ChQuote, ChEOL
ChQuote = Chr(34)
ChEOL   = vbCRLF
ChCR    = vbCR

' Define file access modes
Const ModeReading   = 1
Const ModeWriting   = 2
Const ModeAppending = 8

Function Print(TextString)
'	WScript.Echo         TextString
    WScript.StdOut.Write TextString
End Function

Function Read()
	Read = WScript.StdIn.ReadLine
End Function

Function FolderCreate(FolderName)
	Dim FileSystem
	Set FileSystem = WScript.CreateObject("Scripting.FileSystemObject")

	If Not FileSystem.FolderExists(FolderName) Then
		Call Print("Creating folder: " & ChQuote & FolderName & ChQuote & " ... ")

		Dim IsDone, FolderParent, FolderParent2
		IsDone       = False
		FolderParent = FileSystem.GetParentFolderName(FolderName)

		Do
			Dim NewFolder

			If FileSystem.FolderExists(FolderParent) Then
				FolderParent = FileSystem.GetParentFolderName(FolderName)
			End If

			If FileSystem.FolderExists(FolderParent) Then
				Set NewFolder = FileSystem.CreateFolder(FolderName)
				IsDone = True
			Else
				FolderParent2 = FileSystem.GetParentFolderName(FolderParent)

				if FileSystem.FolderExists(FolderParent2) Then
					Set NewFolder = FileSystem.CreateFolder(FolderParent)
				Else
					FolderParent = FolderParent2
				End If
			End If
		Loop Until IsDOne

		Call Print("Done" & ChEOL)
	End If
End Function

Function FolderDelete(FolderName)
	Dim FileSystem
	Set FileSystem = WScript.CreateObject("Scripting.FileSystemObject")

	If FileSystem.FolderExists(FolderName) Then
		FileSystem.DeleteFolder(FolderName)
	End If
End Function

Function FileTxtAppend(FileName, TextData)
	Dim FileSystem, FileObject
	Set FileSystem = WScript.CreateObject("Scripting.FileSystemObject")
	Set FileObject = FileSystem.OpenTextFile(FileName, ModeAppending, True)

	FileObject.Write(TextData)
	FileObject.Close()
End Function

Function FileBinAppend(FileName, BinaryData)
	Dim FileSystem, FileObject, I, MaxI, TextData
	Set FileSystem = WScript.CreateObject("Scripting.FileSystemObject")
	Set FileObject = FileSystem.OpenTextFile(FileName, ModeAppending, True)

	MaxI     = LenB(BinaryData)
	TextData = ""

	For I = 1 To MaxI
		TextData = TextData & Chr(AscB(MidB(BinaryData, I, 1)))
	Next

	FileObject.Write(TextData)
	FileObject.Close()
End Function

Function FileDelete(FileName)
	Dim FileSystem
	Set FileSystem = WScript.CreateObject("Scripting.FileSystemObject")

	If FileSystem.FileExists(FileName) Then
		FileSystem.DeleteFile(FileName)
	End If
End Function

Function FileDownload(URL, FileName)
	Const ChunkSize          = 10240 '10 KB
	Const PartialContentCode = 206

	Dim Chunk, FileSize
	Chunk    = 0
	FileSize = 0

	Do
		Dim FirstByte, LastByte
		FirstByte = Chunk * ChunkSize
		LastByte  = FirstByte + ChunkSize - 1
		Chunk     = Chunk + 1

		Dim HttpRequest
'		Set HttpRequest = WScript.CreateObject("Microsoft.XMLHTTP")  'msxml3.dll
'		Set HttpRequest = WScript.CreateObject("Msxml2.XMLHttp.6.0") 'msxml6.dll
		Set HttpRequest = WScript.CreateObject("WinHttp.WinHttpRequest.5.1")

		HttpRequest.Open "GET", URL, False
		HttpRequest.SetRequestHeader "Range", "bytes=" & FirstByte & "-" & LastByte
		HttpRequest.Send ""

		Dim Received, Status
		Received = CLng(HttpRequest.GetResponseHeader("Content-Length"))
		Status   = HttpRequest.Status

		If Status <> PartialContentCode Then
			Call Print(ChEOL & "Error" & ChEOL)
			Call FileDelete(FileName)
			Exit Do
		End If

		Call FileBinAppend(FileName, HttpRequest.ResponseBody)

		FileSize = FileSize + Received
		Call Print(ChCR & FileSize & " bytes")
	Loop Until Received <> ChunkSize

	Call Print(ChEOL & "Done" & ChEOL)
End Function

Function WaitEnter()
	Call Print("Press [ENTER] to continue...")
	Call Read()
End Function

Function ForceConsole()
	Dim Interpreter
	Interpreter = "cscript.exe"

	If InStr(LCase(WScript.FullName), Interpreter) = 0 Then
		WShell.Run Interpreter & " //NoLogo " & ChQuote & WScript.ScriptFullName & ChQuote
		WScript.Quit
	End If
End Function

Function CheckBaseFile()
	Dim FileSystem
	Set FileSystem = WScript.CreateObject("Scripting.FileSystemObject")

	If Not FileSystem.FileExists(BaseFile) Then
		Call Print(ChEOL & "Downloading " & ChQuote & UpdateBaseFile & ChQuote & " ... " & ChEOL)
		Call FileDownload(UpdateBaseURL, BaseFile)
		Call Print(ChEOL)
	End If
End Function

Function ListRequiredUpdates()
	Dim FileSystem
	Set FileSystem = WScript.CreateObject("Scripting.FileSystemObject")

	If Not FileSystem.FileExists(DescList) Then
		Call Print(ChEOL & "Connecting to " & ChQuote & UpdateBaseFile & ChQuote & " ... ")
		Dim UpdateManager, UpdateService
		Set UpdateManager = CreateObject("Microsoft.Update.ServiceManager")
		Set UpdateService = UpdateManager.AddScanPackageService("Offline Sync Service", BaseFile, 1)
		' (null): В этом объекте нет подписи.
		Call Print("Done" & ChEOL)

		Call Print("Starting new update session ... ")
		Dim UpdateSession, UpdateSearcher
		Set UpdateSession  = CreateObject("Microsoft.Update.Session")
		Set UpdateSearcher = UpdateSession.CreateUpdateSearcher()
		UpdateSearcher.ServerSelection = 3 'ssOthers
		UpdateSearcher.ServiceID       = UpdateService.ServiceID
		Call Print("Done" & ChEOL)

		Call Print("Searching for non-installed updates ... ")
		Dim SearchResult
		Set SearchResult = UpdateSearcher.Search("IsInstalled=0")
		Call Print("Done" & ChEOL)

'		Dim RequiredUpdates
'		Set RequiredUpdates = SearchResult.Updates

		Call Print("Generating the list of updates ... ")
		Dim I, MaxI
		MaxI = SearchResult.Updates.Count - 1

		For I = 0 to MaxI
			Dim RequiredUpdate
			Set RequiredUpdate = SearchResult.Updates.Item(I)

			Dim UpdateDesc
			UpdateDesc = RequiredUpdate.Title & ChEOL

			Call FileTxtAppend(DescList, UpdateDesc)
		Next
		Call Print("Done" & ChEOL)
	End If
End Function

' Use command line interface
Call ForceConsole()

' Create cache subfolders when it is needed
Call FolderCreate(BaseFolder)
Call FolderCreate(UpdatesFolder)

' Get base file if it is absent
Call CheckBaseFile()

' Generate the list of required updates
Call ListRequiredUpdates()

' End of script
Call WaitEnter()
WScript.Quit
