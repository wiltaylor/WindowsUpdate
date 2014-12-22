'========================================================================================================
' Description: This script uses a wsusscn2.cab file to scan for what updates have not been installed.
'              The script will then output an xml file which can be imported into powershell with
'              Import-CliXml cmdlet.
' Created    : 21/12/2014
' Author     : Wil Taylor (wilfridtaylor@gmail.com) 
'========================================================================================================
Option Explicit

'Will force execution to happen in cscript.
RunInCScript()

If WScript.Arguments.Named.Exists("?") Then Usage()

Dim FSO: Set FSO = CreateObject("Scripting.FileSystemObject")
Dim WshShell : Set WshShell = CreateObject("Wscript.Shell")
Dim cabLocation, reportPath

If Not WScript.Arguments.Named.Exists("CAB") Then 
	cabLocation = FSO.GetAbsolutePathName(".\wsusscn2.cab")
Else
	cabLocation = FSO.GetAbsolutePathName(WScript.Arguments.Named.Item("CAB"))
End If

If Not WScript.Arguments.Named.Exists("Output") Then 
	reportPath = "report.xml"
Else
	reportPath = WScript.Arguments.Named.Item("Output")
End If

Dim Session: Set Session = CreateObject("Microsoft.Update.Session")
Dim Searcher: Set Searcher = Session.CreateUpdateSearcher()
Dim ServiceManager: Set ServiceManager = CreateObject("Microsoft.Update.ServiceManager")
Dim UpdateService: Set UpdateService = ServiceManager.AddScanPackageService("Offline Sync Service", cabLocation)  
Dim Results, u, b, d
Dim XmlDoc: set XmlDoc = CreateObject("MSXML2.DOMDocument")
Dim XmlRoot: Set XmlRoot = XmlDoc.appendChild(XmlDoc.CreateElement("Objs"))
Dim RefID: RefID = 0
Dim TNF1: TNF1 = false
Dim TNF2: TNF2 = false
Dim TNF3: TNF3 = false

XmlRoot.setAttribute "Version" ,"1.1.0.1"
XmlRoot.setAttribute "xmlns" ,"http://schemas.microsoft.com/powershell/2004/04"

Searcher.ServerSelection = 3
Searcher.ServiceID = UpdateService.ServiceID

set Results = Searcher.Search("IsInstalled=0")

Dim Title, Description, SupportURL, RebootRequired, IsUninstallable, Severity, RevisionNumber, UpdateID
Dim UpdateObject, DownloadList, UpdateItem, DLRevision, DLURL, DLName, DLUpdateID
For Each u In Results.Updates
	On Error Resume Next 
	Title = u.title
	Description = u.Description
	SupportURL = u.SupportUrl
	RebootRequired = ""
	RebootRequired = u.RebootRequired
	IsUninstallable = ""
	IsUninstallable = u.IsUninstallable
	RevisionNumber = u.Identity.RevisionNumber
	UpdateID = u.Identity.UpdateID
	On Error Goto 0
		
	Set UpdateObject = AddUpdate(Title, Description, SupportURL, RebootRequired, IsUninstallable, Severity, RevisionNumber, UpdateID)
	Set DownloadList = AddDownloads(UpdateObject)
		
	For Each b in u.BundledUpdates
		For Each d in b.DownloadContents
			
			DLRevision = b.Identity.RevisionNumber
			DLURL = d.DownloadURL
			DLName = b.Title
			DLUpdateID = b.Identity.UpdateID
			
			Set UpdateItem = AddDownloadItem(DLRevision, DLURL, DLName, DLUpdateID)
			DownloadList.appendChild(UpdateItem)
		Next
	Next
Next

XmlDoc.Save reportPath

Function AddUpdate(Title, Description, SupportURL, RebootRequired, IsUninstallable, Severity, RevisionNumber, UpdateID) 
	Dim ElementRef: ElementRef = GetNewRefID()
	
	'<Obj RefId='X'>
	Dim ObjElement: Set ObjElement = XmlDoc.CreateElement("Obj")
	ObjElement.setAttribute "RefId" ,ElementRef
	
	If Not TNF1 Then
		'<TN RefId="0">
		'  <T>System.Management.Automation.PSCustomObject</T>
		'  <T>System.Object</T>
		'</TN>
		Dim TNElement: Set TNElement = XmlDoc.CreateElement("TN")
		Dim TEl1: Set TEl1 = XmlDoc.CreateElement("T")
		Dim TEl2: Set TEl2 = XmlDoc.CreateElement("T")
		TNElement.setAttribute "RefId" ,"0"
		TEl1.Text = "System.Management.Automation.PSCustomObject"
		TEl2.Text = "System.Object"
		TNElement.appendChild Tel1
		TNElement.appendChild Tel2 
		ObjElement.appendChild TNElement
		TNF1 = true
	Else
		Dim TNElement2: Set TNElement2 = XmlDoc.CreateElement("TNRef")
		TNElement2.setAttribute "RefId" ,"0"
		ObjElement.appendChild TNElement2
	End If

    '<MS>
    '  <S N="Title">Intel Corporation - Graphics Adapter WDDM1.1, Graphics Adapter WDDM1.2, Graphics Adapter WDDM1.3 - Intel(R) HD Graphics 4600</S>
    '  <S N="Description">Intel Corporation Graphics Adapter WDDM1.1, Graphics Adapter WDDM1.2, Graphics Adapter WDDM1.3 software update released in September, 2014</S>
    '  <S N="SupportUrl">http://support.microsoft.com/select/?target=hub</S>
    '  <B N="RebootRequired">false</B>
    '  <B N="IsUninstallable">false</B>
	'...
	Dim MSElement: set MSElement = XmlDoc.CreateElement("MS")
	Dim STitle: Set STitle = XmlDoc.CreateElement("S")
	Dim SDescription: Set SDescription = XmlDoc.CreateElement("S")
	Dim SSupportUrl: Set SSupportUrl = XmlDoc.CreateElement("S")
	Dim BRebootRequired: Set BRebootRequired = XmlDoc.CreateElement("B")
	Dim BIsUninstallable: Set BIsUninstallable = XmlDoc.CreateElement("B")
	Dim SSeverity: Set SSeverity = XmlDoc.CreateElement("S")
	Dim I32RevisionNumber: Set I32RevisionNumber = XmlDoc.CreateElement("I32")
	Dim SUpdateID: Set SUpdateID = XmlDoc.CreateElement("S")
	
	STitle.setAttribute "N", "Title"
	SDescription.setAttribute "N", "Description"
	SSupportUrl.setAttribute "N", "SupportUrl"
	BRebootRequired.setAttribute "N", "RebootRequired"
	BIsUninstallable.setAttribute "N", "IsUninstallable"
	SSeverity.setAttribute "N", "Severity"
	I32RevisionNumber.setAttribute "N", "RevisionNumber"
	SUpdateID.setAttribute "N", "UpdateID"
	
	STitle.Text = Title
	SDescription.Text = Description
	SSupportUrl.Text = SupportURL
	BRebootRequired.Text = RebootRequired
	BIsUninstallable.Text = IsUninstallable
	SSeverity.Text = Severity
	I32RevisionNumber.Text = RevisionNumber
	SUpdateID.Text = UpdateID
	
	MSElement.appendChild STitle
	MSElement.appendChild SDescription
	MSElement.appendChild SSupportUrl
	MSElement.appendChild BRebootRequired
	MSElement.appendChild BIsUninstallable
	MSElement.appendChild SSeverity
	MSElement.appendChild I32RevisionNumber
	MSElement.appendChild SUpdateID
	ObjElement.appendChild MSElement
	
	XmlRoot.appendChild ObjElement
	
	Set AddUpdate = MSElement
End Function

Function AddDownloadItem(Revision, URL, Name, UpdateID)
	Dim ElementRef: ElementRef = GetNewRefID()
	
	'<Obj RefID='X'>
	Dim ObjElement: Set ObjElement = XmlDoc.CreateElement("Obj")
	ObjElement.setAttribute "RefId", ElementRef
	
	'<TN RefId="1">
	'  <T>System.Object[]</T>
	'  <T>System.Array</T>
	'  <T>System.Object</T>
	'</TN>
	If Not TNF3 Then
		Dim TNElement: Set TNElement = XmlDoc.CreateElement("TN")
		Dim TEl1: Set TEl1 = XmlDoc.CreateElement("T")
		Dim TEl2: Set TEl2 = XmlDoc.CreateElement("T")
		TNElement.setAttribute "RefId" ,"2"
		TEl1.Text = "System.Collections.Hashtable"
		TEl2.Text = "System.Object"
		TNElement.appendChild TEl1
		TNElement.appendChild TEl2
		ObjElement.appendChild TNElement
		TNF3 = true
	Else
		Dim TNElement2: Set TNElement2 = XmlDoc.CreateElement("TNRef")
		TNElement2.setAttribute "RefId" ,"2"
		ObjElement.appendChild(TNElement2)
	End If
	
	'<DCT>
	'  <En>
	'	<S N="Key">Revision</S>
	'	<I32 N="Value">205</I32>
	'  </En>
	'  <En>
	'	<S N="Key">URL</S>
	'	<S N="Value">http://download.windowsupdate.com/c/msdownload/update/software/updt/2014/11/windows8.1-kb3000850-x64_6587d8609636fd5abd2c5659b805f754810b6c96.cab</S>
	'  </En>
	'  <En>
	'	<S N="Key">Name</S>
	'	<S N="Value">windows8.1-kb3000850-x64</S>
	'  </En>
	'  <En>
	'	<S N="Key">UpdateID</S>
	'	<S N="Value">f3374921-0ea9-4939-bcc3-4a67c677c84e</S>
	'  </En>
	'</DCT>
	Dim DCTElement: Set DCTElement = XmlDoc.CreateElement("DCT")
	Dim En1: Set En1 = XmlDoc.CreateElement("En")
	Dim En2: Set En2 = XmlDoc.CreateElement("En")
	Dim En3: Set En3 = XmlDoc.CreateElement("En")
	Dim En4: Set En4 = XmlDoc.CreateElement("En")
	Dim Key1: set Key1 = XmlDoc.CreateElement("S")
	Dim Key2: set Key2 = XmlDoc.CreateElement("S")
	Dim Key3: set Key3 = XmlDoc.CreateElement("S")
	Dim Key4: set Key4 = XmlDoc.CreateElement("S")
	Dim Value1: set Value1 = XmlDoc.CreateElement("I32")
	Dim Value2: set Value2 = XmlDoc.CreateElement("S")
	Dim Value3: set Value3 = XmlDoc.CreateElement("S")
	Dim Value4: set Value4 = XmlDoc.CreateElement("S")
	Key1.setAttribute "N", "Key"
	Key2.setAttribute "N", "Key"
	Key3.setAttribute "N", "Key"
	Key4.setAttribute "N", "Key"
	
	Key1.Text = "Revision"
	Key2.Text = "URL"
	Key3.Text = "Name"
	Key4.Text = "UpdateID"
	
	Value1.setAttribute "N", "Value"
	Value2.setAttribute "N", "Value"
	Value3.setAttribute "N", "Value"
	Value4.setAttribute "N", "Value"
	
	Value1.Text = Revision 
	Value2.Text = URL
	Value3.Text = Name
	Value4.Text = UpdateID
	
	En1.appendChild Key1
	En1.appendChild Value1
	En2.appendChild Key2
	En2.appendChild Value2
	En3.appendChild Key3
	En3.appendChild Value3
	En4.appendChild Key4
	En4.appendChild Value4
	
	DCTElement.appendChild En1
	DCTElement.appendChild En2
	DCTElement.appendChild En3
	DCTElement.appendChild En4
	ObjElement.appendChild DCTElement
	
	Set AddDownloadItem = ObjElement
	
End Function

Function AddDownloads(UpdateNode) 
	Dim ElementRef: ElementRef = GetNewRefID()
	
	'<Obj RefId='X'>
	Dim ObjElement: Set ObjElement = XmlDoc.CreateElement("Obj")
	ObjElement.setAttribute "RefId", ElementRef
	ObjElement.setAttribute "N", "Downloads"
	
	'<TN RefId="1">
	'  <T>System.Object[]</T>
	'  <T>System.Array</T>
	'  <T>System.Object</T>
	'</TN>
	If Not TNF2 Then
		Dim TNElement: Set TNElement = XmlDoc.CreateElement("TN")
		Dim TEl1: Set TEl1 = XmlDoc.CreateElement("T")
		Dim TEl2: Set TEl2 = XmlDoc.CreateElement("T")
		Dim TEl3: Set TEl3 = XmlDoc.CreateElement("T")
		TNElement.setAttribute "RefId" ,"1"
		TEl1.Text = "System.Object[]"
		TEl2.Text = "System.Array"
		TEl3.Text = "System.Object"
		TNElement.appendChild TEl1
		TNElement.appendChild TEl2
		TNElement.appendChild TEl3
		ObjElement.appendChild TNElement
		TNF2 = true
	Else
		Dim TNElement2: Set TNElement2 = XmlDoc.CreateElement("TNRef")
		TNElement2.setAttribute "RefId" ,"1"
		ObjElement.appendChild TNElement2
	End If
	
	'<LST>
	Dim LSTElement: Set LSTElement = XmlDoc.CreateElement("Obj")
	ObjElement.appendChild LSTElement
	UpdateNode.appendChild ObjElement
	
	set AddDownloads = LSTElement
End Function

Function GetNewRefID()
	GetNewRefID = RefID
	RefID = RefID + 1
End Function

public Sub Usage()
	WScript.echo "Usage"
	WScript.echo "cscript.exe ScanUpdate.vbs /CAB:(CAB) /OUTPUT:(OUTPUT)"
	WScript.echo ""
	WScript.echo " CAB    - Path to wsusscn2.cab"
	WScript.echo " OUTPUT - Path to location to save xml report to."
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
