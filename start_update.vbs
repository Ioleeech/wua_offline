Dim WShell
Set WShell = WScript.CreateObject("WScript.Shell")

Dim FileSystem
Set FileSystem = WScript.CreateObject("Scripting.FileSystemObject")

' Windows Update Agent (WUA) base file for offline updates
Dim UpdateBaseURL, UpdateBaseFile, UpdateDescList
UpdateBaseURL  = "http://go.microsoft.com/fwlink/p/?LinkID=74689"
UpdateBaseFile = "wsusscn2.cab"
UpdateDescList = "updates.txt"

' Define the tree of subfolders
Dim CurrentPath, CacheFolder, BaseFile, DescList
CurrentPath = WShell.CurrentDirectory
CacheFolder = CurrentPath & "\cache"
BaseFile    = CacheFolder & "\" & UpdateBaseFile
DescList    = CacheFolder & "\" & UpdateDescList

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
	If FileSystem.FolderExists(FolderName) Then
		FileSystem.DeleteFolder(FolderName)
	End If
End Function

Function FileTxtAppend(FileName, TextData)
	Dim FileObject
	Set FileObject = FileSystem.OpenTextFile(FileName, ModeAppending, True)

	FileObject.Write(TextData)
	FileObject.Close()
End Function

Function FileBinAppend(FileName, BinaryData)
	Dim FileObject, I, MaxI, TextData
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
	If Not FileSystem.FileExists(BaseFile) Then
		Call Print(ChEOL & "Downloading " & ChQuote & UpdateBaseFile & ChQuote & " ... " & ChEOL)
		Call FileDownload(UpdateBaseURL, BaseFile)
		Call Print(ChEOL)
	End If
End Function

Function DescListAddURL(ByRef DownloadContents)
	Dim I, MaxI
	MaxI = DownloadContents.Count - 1

	For I = 0 To MaxI
		Dim ContentInfo
		Set ContentInfo = DownloadContents.Item(I)

		Call FileTxtAppend(DescList, "URL      : " & ContentInfo.DownloadUrl & ChEOL)
	Next
End Function

Function DescListAddBundled(ByRef BundledUpdates)
	Dim I, MaxI
	MaxI = BundledUpdates.Count - 1

	For I = 0 To MaxI
		Dim DownloadContents
		Set DownloadContents = BundledUpdates.Item(I).DownloadContents

		Call DescListAddURL(DownloadContents)
	Next
End Function

Function ListRequiredUpdates(ByRef SearchResult)
	Const RebootNever          = 0
	Const RebootAlwaysRequires = 1
	Const RebootCanRequest     = 2

	Dim RequiredUpdates
	Set RequiredUpdates = SearchResult.Updates

	Call FileDelete(DescList)
	Call FileTxtAppend(DescList, RequiredUpdates.Count & " updates in list" & ChEOL)

	Dim I, MaxI
	MaxI = RequiredUpdates.Count - 1

	For I = 0 To MaxI
		Dim UpdateInfo
		Set UpdateInfo = RequiredUpdates.Item(I)

		Call FileTxtAppend(DescList, ChEOL & "Title    : " & UpdateInfo.Title & ChEOL)

		Select Case UpdateInfo.InstallationBehavior.RebootBehavior
			Case RebootNever
				Call FileTxtAppend(DescList, "Reboot   : No"         & ChEOL)
			Case RebootAlwaysRequires
				Call FileTxtAppend(DescList, "Reboot   : Required"   & ChEOL)
			Case RebootCanRequest
				Call FileTxtAppend(DescList, "Reboot   : Recomended" & ChEOL)
			Case Else
				Call FileTxtAppend(DescList, "Reboot   : Unknown"    & ChEOL)
		End Select

		If UpdateInfo.IsDownloaded = True Then
			Call FileTxtAppend(DescList, "Download : Done"     & ChEOL)
		Else
			Call FileTxtAppend(DescList, "Download : Required" & ChEOL)

			Dim DownloadContents, BundledUpdates
			Set DownloadContents = UpdateInfo.DownloadContents
			Set BundledUpdates   = UpdateInfo.BundledUpdates

			If DownloadContents.Count <> 0 Then
				Call DescListAddURL(DownloadContents)
			ElseIf BundledUpdates.Count <> 0 Then
				Call DescListAddBundled(BundledUpdates)
			Else
				Call FileTxtAppend(DescList, "URL      : None" & ChEOL)
			End If
		End If
	Next
End Function

Class DummyClass
	Public Default Function DummyFunction()
	End Function		
End Class

Function DownloadRequiredUpdates(ByRef UpdateDownloader, ByRef SearchResult)
	UpdateDownloader.Updates = SearchResult.Updates
'	UpdateDownloader.Download()

	Dim DummyDictionary
	Set DummyDictionary = WScript.CreateObject("Scripting.Dictionary")
	Call DummyDictionary.Add("DummyFunction", New DummyClass)

	Dim DownloadJob
	Set DownloadJob = UpdateDownloader.BeginDownload(DummyDictionary.Item("DummyFunction"), DummyDictionary.Item("DummyFunction"), vbNull)

	Dim I, MaxI, Percent, Completed
	MaxI = DownloadJob.Updates.Count

	Do
		' 500 ms timeout
		WScript.Sleep 500

		Completed = DownloadJob.IsCompleted
		I         = DownloadJob.GetProgress.CurrentUpdateIndex + 1
		Percent   = DownloadJob.GetProgress.PercentComplete

		Call Print(ChCR & "Update " & I & "/" & MaxI & " (" & Percent & "%)  ")
	Loop Until Completed = True

	Call Print(ChEOL)
End Function

Function InstallRequiredUpdates(ByRef UpdateInstaller, ByRef SearchResult)
	Dim DownloadedUpdates
	Set DownloadedUpdates = WScript.CreateObject("Microsoft.Update.UpdateColl")

	Dim I, MaxI, Completed
	MaxI = SearchResult.Updates.Count - 1

	For I = 0 To MaxI
		Dim UpdateInfo
		Set UpdateInfo = SearchResult.Updates.Item(I)

		If UpdateInfo.IsDownloaded = True Then
			DownloadedUpdates.Add(UpdateInfo)	
		End If
	Next

	UpdateInstaller.Updates = DownloadedUpdates
'	UpdateInstaller.Install()

	Dim DummyDictionary
	Set DummyDictionary = WScript.CreateObject("Scripting.Dictionary")
	Call DummyDictionary.Add("DummyFunction", New DummyClass)

	Dim InstallJob
	Set InstallJob = UpdateInstaller.BeginInstall(DummyDictionary.Item("DummyFunction"), DummyDictionary.Item("DummyFunction"), vbNull)

	MaxI = InstallJob.Updates.Count

	Do
		' 500 ms timeout
		WScript.Sleep 500

		Completed = InstallJob.IsCompleted
		I         = InstallJob.GetProgress.CurrentUpdateIndex + 1
		Percent   = InstallJob.GetProgress.CurrentUpdatePercentComplete

		Call Print(ChCR & "Update " & I & "/" & MaxI & " (" & Percent & "%)  ")
	Loop Until Completed = True

	Call Print(ChEOL)
End Function

Function GetRequiredUpdates()
	Const SelectionDefault       = 0
	Const SelectionManagedServer = 1
	Const SelectionWindowsUpdate = 2
	Const SelectionOthers        = 3

	Call Print("Registering " & ChQuote & UpdateBaseFile & ChQuote & " package ... ")
	Dim UpdateManager, UpdateService
	Set UpdateManager = WScript.CreateObject("Microsoft.Update.ServiceManager")
	Set UpdateService = UpdateManager.AddScanPackageService("Offline Sync Service", BaseFile)

	Dim UpdateSession, UpdateSearcher
	Set UpdateSession  = WScript.CreateObject("Microsoft.Update.Session")
	Set UpdateSearcher = UpdateSession.CreateUpdateSearcher()
	UpdateSearcher.ServerSelection = SelectionOthers
	UpdateSearcher.ServiceID       = UpdateService.ServiceID
	Call Print("Done" & ChEOL)

	Call Print("Searching for non-installed updates ... ")
	Dim SearchResult
'	Set SearchResult = UpdateSearcher.Search("IsInstalled=0 and Type='Software'")
	Set SearchResult = UpdateSearcher.Search("IsInstalled=0")
	Call Print("Done" & ChEOL)

	If SearchResult.Updates.Count <> 0 Then
		Call Print("Generating the list of updates ... ")
		Call ListRequiredUpdates(SearchResult)
		Call Print("Done" & ChEOL & ChEOL)

		Call Print("Downloading non-installed updates ... " & ChEOL)
		Dim UpdateDownloader
		Set UpdateDownloader = UpdateSession.CreateUpdateDownloader() 
		Call DownloadRequiredUpdates(UpdateDownloader, SearchResult)
		Call Print("Done" & ChEOL & ChEOL)

		Call Print("Installing updates ... " & ChEOL)
		Dim UpdateInstaller
		Set UpdateInstaller = UpdateSession.CreateUpdateInstaller()
		Call InstallRequiredUpdates(UpdateInstaller, SearchResult)
		Call Print("Done" & ChEOL & ChEOL)
	Else
		Call Print("All updates are already installed, there is nothing to do!" & ChEOL)
		Call FileDelete(DescList)
	End If

	Call Print("Unegistering " & ChQuote & UpdateBaseFile & ChQuote & " package ... ")
	UpdateManager.RemoveService(UpdateService.ServiceID)
	Call Print("Done" & ChEOL)
End Function

' Use command line interface
Call ForceConsole()

' Create cache subfolder when it is needed
Call FolderCreate(CacheFolder)

' Get last base file when it is needed
Call CheckBaseFile()

' Get required updates
Call GetRequiredUpdates()

' End of script
Call WaitEnter()
WScript.Quit
