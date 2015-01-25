function Receive-ScanCab {
    <# 
 .Synopsis
  Downloads the lastest version of wsusscan2.cab.

 .Description
  Downloads the lastest version of wsusscan2.cab for offline security patch scanning.
 
 .Parameter Path
  Path to save wsusscan2.cab. If ommited it will download to the current directory.

 .Parameter LegacyTools
  This will copy InstallUpdates.vbs and ScanUpdate.vbs into the same directory as the wsusscan2.cab file
  for use on machine which don't have powershell 4 installed.

 .Example
  #Downloads wsusscan2.cab to the current directory.
  Receive-ScanCab

 .Example
  Receive-ScanCab -Path "c:\temp\wsusscn2.cab" -LegacyTools

 .LINK
  http://www.win32.io/cmdlet/Receive-ScanCab.html

 .LINK
  about_WindowsUpdateModule
#>
    param($Path = ".\wsusscn2.cab", [switch]$LegacyTools)    
    
    Invoke-WebRequest "go.microsoft.com/fwlink/?LinkId=76054" -OutFile $Path
    $path = Resolve-Path $Path 

    if($LegacyTools) {
        $folder = Split-Path $Path -Parent
        Copy-Item $PSScriptRoot\LegacyScripts\InstallUpdates.vbs -Destination $folder\InstallUpdates.vbs
        Copy-Item $PSScriptRoot\LegacyScripts\ScanUpdate.vbs -Destination $folder\ScanUpdate.vbs
    }
    
}

Export-ModuleMember Receive-ScanCab