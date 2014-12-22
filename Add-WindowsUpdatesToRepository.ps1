function Add-WindowsUpdatesToRepository {
<# 
 .Synopsis
  Downloads updates from a download list.

 .Description
  Downloads updates from a download list created with the Get-WindowsUpdateDownloadList cmdlet or legacy ScanUpdate.vbs script.
 
 .Parameter UpdateList
  Array of WindowsDownloadList objects download.

 .Parameter Path
  Directory in which to store updates.

 .Inputs
  WindowsDownloadList objects. Created by Get-WindowsUpdateDownloadList or ScanUpdate.vbs.

 .Outputs
  Nothing

 .Example
  Add-WindowsUpdatesToRepository -UpdateList $updatelist -Path c:\temp\updates

 .LINK
  about_WindowsUpdateModule
#>
    [CmdletBinding ()]
    param([Parameter(Mandatory=$True,ValueFromPipeline=$True)]$UpdateList, $Path = ".\UpdateRepository") 

    if(-not(Test-Path $Path)) { New-item $path -ItemType Directory | Out-Null }

    foreach($update in $UpdateList) {
        foreach($d in $update.Downloads) {
            $filename = $d["URL"].SubString( $d["URL"].LastIndexOf("/") + 1)

            if(-not(Test-Path $Path\$filename)) {
                #Invoke-WebRequest $d["URL"] -OutFile $Path\$filename

                Write-Verbose "Downloading: $($d["URL"])"

                $client = New-Object System.Net.WebClient
                $client.DownloadFile($d["URL"], "$Path\$filename")
            } else {
                Write-Verbose "Skipping: $($d["URL"]) - Already exists in file system."
            }
        }
    }
}

Export-ModuleMember Add-WindowsUpdatesToRepository