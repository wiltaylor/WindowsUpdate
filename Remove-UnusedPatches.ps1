﻿function Remove-UnusedPatches {
<# 
 .Synopsis
  Removes unused patches in a update repository folder.

 .Description
  Removes unused patches in a update repository folder. Simple pass it an updatelist generated by Get-WindowsUpdateDownloadList

 .Parameter Path
  Path to repository to clean up.

 .Inputs
  Update List objects to scan for required files.
 
 .Example
  #$Updates contains a download list object array.
  $updates | Remove-UnusedPatches -Path c:\PatchRepository

   .Example
  #Shows the user what will be deleted if this command is run.
  $updates | Remove-UnusedPatches -Path c:\PatchRepository -WhatIf

   .LINK
  about_WindowsUpdateModule
#>
    [CmdletBinding(
        SupportsShouldProcess=$true,
        ConfirmImpact="Low")]
    param([Parameter(Mandatory=$True,ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$false)]$Objects, [Parameter(Mandatory=$True)]$Path)

    BEGIN {
        $filelist = Get-ChildItem $Path | ForEach-Object  {$_.FullName.ToLower() }
    }

    PROCESS {   
        foreach($update in $Objects) {
            foreach($download in $update.Downloads) {
                $filename = $download["URL"].SubString( $download["URL"].LastIndexOf("/") + 1)
                $filename = "$Path\$filename"
                $filename = $filename.ToLower()
                $filelist = $filelist -ne $filename

                Write-Verbose "Keeping $filename"
            }
        }
    }

    END {
        if($filelist.count -lt 1) { return }
        Remove-Item $filelist -Force
    }

 
}

Export-ModuleMember Remove-UnusedPatches