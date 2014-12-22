'========================================================================================================
' Description: Installs Windows Updates from an offline repository.
' Created    : 21/12/2014
' Author     : Wil Taylor (wilfridtaylor@gmail.com) 
'========================================================================================================
Option Explicit

'Will force execution to happen in cscript.
RunInCScript()

If WScript.Arguments.Named.Exists("?") Then Usage()

Dim FSO: Set FSO = CreateObject("Scripting.FileSystemObject")
Dim WshShell : Set WshShell = CreateObject("Wscript.Shell")
Dim cabLocation, repositoryPath

If Not WScript.Arguments.Named.Exists("CAB") Then 
	cabLocation = FSO.GetAbsolutePathName(".\wsusscn2.cab")
Else
	cabLocation = FSO.GetAbsolutePathName(WScript.Arguments.Named.Item("CAB"))
End If

If Not WScript.Arguments.Named.Exists("REPOSITORY") Then 
	repositoryPath = FSO.GetAbsolutePathName(".\UpdateRepository")
Else
	repositoryPath = FSO.GetAbsolutePathName(WScript.Arguments.Named.Item("REPOSITORY"))
End If

Dim UpdateCollection: set UpdateCollection = CreateObject("Microsoft.Update.UpdateColl")
Dim Session: Set Session = CreateObject("Microsoft.Update.Session")
Dim Searcher: Set Searcher = Session.CreateUpdateSearcher()
Dim ServiceManager: Set ServiceManager = CreateObject("Microsoft.Update.ServiceManager")
Dim UpdateService: Set UpdateService = ServiceManager.AddScanPackageService("Offline Sync Service", cabLocation)  
Dim Installer: Set Installer = Session.CreateUpdateInstaller
Dim update, bundle, download, filename, Results, strCol
Searcher.ServerSelection = 3
Searcher.ServiceID = UpdateService.ServiceID

set Results = Searcher.Search("IsInstalled=0")

For Each update in Results.Updates
	For Each bundle in update.BundledUpdates
		For Each download in bundle.DownloadContents
			filename = Mid(download.DownloadUrl, InStrRev(download.DownloadUrl, "/") + 1)
			
			If FSO.FileExists(repositoryPath & "\" & filename) Then
				set strCol = CreateObject("Microsoft.Update.StringColl")
				strCol.Add(repositoryPath & "\" & filename)
				bundle.CopyToCache strCol
			End If
		Next
	Next
Next

set Results = Searcher.Search("IsInstalled=0")

For Each update in Results.Updates
	If update.IsDownloaded Then
		UpdateCollection.Add update
		
		If Not update.EulaAccepted Then update.AcceptEula
		
	End If
Next

Installer.Updates = UpdateCollection
Installer.Install

public Sub Usage()
	WScript.echo "Usage"
	WScript.echo "cscript.exe InstallUpdates.vbs /CAB:(CAB) /REPOSITORY:(OUTPUT)"
	WScript.echo ""
	WScript.echo " CAB        - Path to wsusscn2.cab"
	WScript.echo " REPOSITORY - Path to folder containing cached windows updates."
	Wscript.quit
End Sub

Sub RunInCScript()
    Dim Arg, Str
    If Not LCase( Right( WScript.FullName, 12 ) ) = "\cscript.exe" Then
        For Each Arg In WScript.Arguments
            If InStr( Arg, " " ) Then Arg = """" & Arg & """"
            Str = Str & " " & Arg
        Next
		
        WScript.Quit(CreateObject( "WScript.Shell" ).Run( _
            "cscript //nologo """ & _
            WScript.ScriptFullName & _
            """ " & Str))
    End If
End Sub