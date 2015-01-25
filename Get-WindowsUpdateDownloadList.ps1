function Get-WindowsUpdateDownloadList {
    <# 
 .Synopsis
  Creates an array of WindowsUpdateDownloadList objects.

 .Description
  Creates an array of WindowsUpdateDownloadList objects. These objects are used by Add-WindowsUpdatesToRepository to 
  download required updates for use offline.
 
 .Inputs
  Update objects retrived by Get-WindowsUpdate.

 .Outputs
  Returns WindowsUpdateDownloadList objects for Add-WindowsUpdatesToRepository.

 .Notes
  You may use ScanUpdate.vbs to generate compatable xml files that can be used by Import-CliXml to also generate 
  WindowsUpdateDownloadList objects that can be used by Add-WindowsUpdatesToRepository.
   
 .Example
  $downloads = Get-WindowsUpdate -Available | Get-WindowsUpdateDownloadList 

 .LINK
  about_WindowsUpdateModule

 .LINK
  http://www.win32.io/cmdlet/Get-WindowsUpdateDownloadList.html
#>
    [CmdletBinding ()]
    param([Parameter(Mandatory=$True,ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$false)]$Update)

    PROCESS {
        $Obj = New-Object PSObject
        $Obj | Add-Member -MemberType NoteProperty -Name Title -Value $Update.Title
        $Obj | Add-Member -MemberType NoteProperty -Name Description -Value $Update.Description
        $Obj | Add-Member -MemberType NoteProperty -Name SupportUrl -Value $Update.SupportUrl
        $Obj | Add-Member -MemberType NoteProperty -Name RebootRequired -Value $Update.RebootRequired
        $Obj | Add-Member -MemberType NoteProperty -Name IsUninstallable -Value $Update.IsUninstallable
        $Obj | Add-Member -MemberType NoteProperty -Name Severity -Value $Update.MsrcSeverity
        $Obj | Add-Member -MemberType NoteProperty -Name RevisionNumber -Value $Update.Identity.RevisionNumber
        $Obj | Add-Member -MemberType NoteProperty -Name UpdateID -Value $Update.Identity.UpdateID
    
        $Obj | Add-Member -MemberType NoteProperty -Name BulletinIDs -Value @()
        $Obj | Add-Member -MemberType NoteProperty -Name KBArticleIDs -Value @()
        $obj | Add-Member -MemberType NoteProperty -Name Downloads -Value @()
        $obj | Add-Member -MemberType NoteProperty -Name Superseded -Value @()

        foreach($kb in $Update.KBArticleIDs) {$Obj.KBArticleIDs += $kb }
        foreach($b in $Update.SecurityBulletinIDs) {$Obj.BulletinIDs += $b }
        foreach($s in $Update.SupersededUpdateIDs) {$Obj.Superseded += $s }

        foreach($bundle in $Update.BundledUpdates) {
            foreach($d in $bundle.DownloadContents) { 
                $Obj.Downloads += @{Name = $bundle.Title; URL = $d.DownloadURL; Revision = $bundle.Identity.RevisionNumber; UpdateID =  $bundle.Identity.UpdateID}
            }
        }

        $Obj
    }
}

Export-ModuleMember Get-WindowsUpdateDownloadList

