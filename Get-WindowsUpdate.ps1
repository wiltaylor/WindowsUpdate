function Get-WindowsUpdate {
<# 
 .Synopsis
  Gets windows update objects.

 .Description
  Gets windows update objects for use with other cmdlets.
 
 .Parameter Cab
  Path to scan cab. This allows searching for patches while not on the internet.

 .Parameter History
  Returns patch history objects.

 .Parameter Available
  Returns available patches. Uses IsInstalled=0 Query.

 .Parameter Installed
  Returns all installed patches. Uses IsInstalled=1 Query.

 .Parameter Query
  Shows patches returned by the supplied Query.
  Please note you can only search on a limited set of items.
  See http://msdn.microsoft.com/en-us/library/windows/desktop/aa386526(v=vs.85).aspx for more details.

 .Parameter Superseded
  Returns all the Superseded patches that are currently installed.
  
 .Parameter ID
  Returns patch by its GUID. This is not the KB article ID. For that use KB

 .Parameter Revision
  Returns a patch object for the target revision. Use this with ID.

 .Parameter Severity
  Filters results by the Severity. Can be set to None, Critical, Important, Moderate or Low.

 .Parameter Bulletin
  Filters results by Microsoft Security Bulletin ID. e.g. MS14-036

 .Parameter KB
  Filters results by KB Article. Don't append KB to the parameter, just use the number.

 .Parameter Category
  Filters results by Category. Can be set to Application, Connectors, CriticalUpdates, DefinitionUpdates, 
    DeveloperKits, FeaturePacks, Guidance, SecurityUpdates, ServicePacks, Tools, UpdateRollups or Updates
  
 .Example
  #Shows installed updates.
  Get-WindowsUpdate -Installed

 .Example
  #Shows all available updates
  Get-WindowsUpdate -Available

 .Example
  #Shows all updates.
  Get-WindowsUpdate

 .Example
  #Shows all updates by custom Windows Update Search query.
  Get-WindowsUpdate -Query "IsHidden=1"

 .Example
  #Shows all installed superseded updates.
  Get-WindowsUpdate -Superseded

 .Example
  #Gets all updates with ID of 0465b24d-25a0-400b-bf74-d3e0100d8f22 (usual 1 but can be multiple if there are multiple revisions).
  Get-WindowsUpdate -ID 0465b24d-25a0-400b-bf74-d3e0100d8f22

 .Example
  #Gets all updates with ID of 0465b24d-25a0-400b-bf74-d3e0100d8f22 and Revision 201.
  Get-WindowsUpdate -ID 0465b24d-25a0-400b-bf74-d3e0100d8f22 -Revision 201

 .Example
  #Shows all installed superseded updates.
  Get-WindowsUpdate -Severity Critical

 .Example
  #Shows all updates assgined to the Bulletin MS14-036
  Get-WindowsUpdate -Bulletin MS14-036

 .Example
  #Shows all updates assgined to KB2964718
  Get-WindowsUpdate -KB 2964718

 .LINK
  about_WindowsUpdateModule

 .LINK
  http://www.win32.io/cmdlet/Get-WindowsUpdate.html
#>
    param($cab = "", [switch]$History, [switch]$Available, [switch]$Installed ,[string]$Query = "", [switch]$Superseded, [string]$ID="", [string]$Revision="", 
    [ValidateSet("", "None", "Critical", "Important", "Moderate", "Low")]
    [string]$Severity="", [string]$Bulletin="", [string]$KB="", 
    [ValidateSet("", "Application", "Connectors", "CriticalUpdates", "DefinitionUpdates", "DeveloperKits", "FeaturePacks", "Guidance", "SecurityUpdates", "ServicePacks",
    "Tools", "UpdateRollups", "Updates")]
    [string]$Category="")

    if($Revision -ne "" -and $ID -eq "") {
        Write-Error "You must supply ID if you use Revision"
        return
    }

    $Session = New-Object -ComObject Microsoft.Update.Session
    $Searcher = $Session.CreateUpdateSearcher()
   
    if($History) {
        $HistoryCount = $Searcher.GetTotalHistoryCount()
        $Searcher.QueryHistory(1,$HistoryCount)
        return
    }


    if($Superseded) {
        $SupersededList = @()
        $ReturnList = @()

        $UpdateServiceManager = New-Object -ComObject Microsoft.Update.ServiceManager
        if($cab -ne "") {
            $UpdateService = $UpdateServiceManager.AddScanPackageService("Offline Sync Service", $cab)        
            $Searcher.ServerSelection = 3
            $Searcher.ServiceID = $UpdateService.ServiceID
        }

        $results = $Searcher.Search("IsInstalled=1").Updates   

        foreach($u in $results) {
            foreach($s in $u.SupersededUpdateIDs) {
                $SupersededList += $s
            }
        }

        foreach($u in $results) { 
            if ($SupersededList -contains $u.Identity.UpdateID) {$u}

            foreach($b in $u.BundledUpdates) {
                if ($SupersededList -contains $b.Identity.UpdateID) {$u}
            }
        }

        return
    }


    if($Query -eq "") {
        if($Available) { $Query = "IsInstalled=0"}
        if($Installed) { $Query = "IsInstalled=1"}  
        if($ID -ne "" -and $Revision -eq "") { $Query = "UpdateID='$ID'"}
        if($ID -ne "" -and $Revision -ne "") { $Query = "UpdateID='$ID' and RevisionNumber = $Revision"}
    }

    if($Query -eq "") {
        $Query = "IsInstalled=0 or IsInstalled=1"
    }



    if($Query -ne "") {
        $UpdateServiceManager = New-Object -ComObject Microsoft.Update.ServiceManager
        if($cab -ne "") {
            $UpdateService = $UpdateServiceManager.AddScanPackageService("Offline Sync Service", $cab)        
            $Searcher.ServerSelection = 3
            $Searcher.ServiceID = $UpdateService.ServiceID
        }

        $returnData = $Searcher.Search($Query).Updates
    }


    if($Severity -ne ""){
        if($Severity -eq "None") {
            $returnData = $returnData | where MsrcSeverity -eq $null
        } else {
            $returnData = $returnData | where MsrcSeverity -eq $Severity
        }
    }

    if($Bulletin -ne ""){
        $ReturnData = $returnData | where SecurityBulletinIDs -Contains $Bulletin
    }

    if($KB -ne ""){
        $ReturnData = $returnData | where KBArticleIDs -Contains $KB
    }

    if($KB -ne ""){
        $ReturnData = $returnData | where KBArticleIDs -Contains $KB
    }

    if($Category -ne "") {
        
        switch($Category) {
            "Application"       { $guid = "5C9376AB-8CE6-464A-B136-22113DD69801"}
            "Connectors"        { $guid = "434DE588-ED14-48F5-8EED-A15E09A991F6"}
            "CriticalUpdates"   { $guid = "E6CF1350-C01B-414D-A61F-263D14D133B4"}
            "DefinitionUpdates" { $guid = "E0789628-CE08-4437-BE74-2495B842F43B"}
            "DeveloperKits"     { $guid = "E140075D-8433-45C3-AD87-E72345B36078"}
            "FeaturePacks"      { $guid = "B54E7D24-7ADD-428F-8B75-90A396FA584F"}
            "Guidance"          { $guid = "9511D615-35B2-47BB-927F-F73D8E9260BB"}
            "SecurityUpdates"   { $guid = "0FA1201D-4330-4FA8-8AE9-B877473B6441"}
            "ServicePacks"      { $guid = "68C5B0A3-D1A6-4553-AE49-01D3A7827828"}
            "Tools"             { $guid = "B4832BD8-E735-4761-8DAF-37F882276DAB"}
            "UpdateRollups"     { $guid = "28BC880E-0592-4CBF-8F95-C79B17911D5F"}
            "Updates"           { $guid = "CD5FFD1E-E932-4E3A-BF74-18BF0B1BBD83"}
            default             { $guid = ""}
        }

        if($guid -ne "") {
            $returnData = $returnData | where {
                foreach($i in $_.Categories) {
                    if($i.CategoryID.ToUpper() -eq $guid) { return $true }
                }

                $false
            }
        }

    }

    $returnData
}

Export-ModuleMember Get-WindowsUpdate