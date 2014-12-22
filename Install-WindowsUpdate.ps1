function Install-WindowsUpdate {
<# 
 .Synopsis
  Installs Windows Updates.

 .Description
  Installs Windows Updates piped into this command.
 
 .Inputs
  Update objects to install. Use Get-WindowsUpdate to get these objects.

 .Outputs
  Nothing

 .Parameter DownloadOnly
  Only downloads windows updates to windows update cache.

 .Parameter OfflineRepository
  Path to offline repository to retrive patches from. If left blank this will go to Windows Update.
 
 .Example
  #Installs all required windows updates.
  Get-WindowsUpdate -Available | Install-WindowsUpdate

 .LINK
  about_WindowsUpdateModule
#>
    [CmdletBinding ()]
    param([Parameter(Mandatory=$True,ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$false)]$update, [switch]$DownloadOnly, [string]$OfflineRepository = "")

    BEGIN{
        $UpdateCollection = New-Object -ComObject Microsoft.Update.UpdateColl
        $AlreadyDownloaded = New-Object -ComObject Microsoft.Update.UpdateColl
    }

    PROCESS {
        if($update.IsDownloaded)
        {
            $AlreadyDownloaded.Add($update) | Out-Null
        } else {
            $UpdateCollection.Add($update) | Out-Null
        }
    }

    END {
        $Session = New-Object -ComObject Microsoft.Update.Session

        if($OfflineRepository -eq "") {
            $downloader = $Session.CreateUpdateDownloader() 
            $downloader.Updates = $UpdateCollection
            $downloader.Download() | Out-Null
        }else {
            foreach($u in $UpdateCollection) {
                foreach($b in $u.BundledUpdates) {
                    foreach($d in $b.DownloadContents) {
                        $filename = $d.DownloadUrl.SubString( $d.DownloadUrl.LastIndexOf("/") + 1)

                        if(Test-Path $OfflineRepository\$filename) {
                            $b.CopyToCache("$OfflineRepository\$filename")
                        }
                    }
                }
            }
        }

        $updatesToInstall = New-Object -ComObject Microsoft.Update.UpdateColl

        foreach($i in $UpdateCollection) {
            if($i.IsDownloaded) { 
                $updatesToInstall.Add($i) |Out-Null
            } else {
                Write-Warning "Failed to download $($i.Title)"
            }
        }

        foreach($i in $AlreadyDownloaded) {
            if($i.IsDownloaded) { 
                $updatesToInstall.Add($i) |Out-Null

                if(-not $i.EulaAccepted) {
                    $i.AcceptEula()
                }

            } else {
                Write-Warning "Failed to download $($i.Title)"
            }
        }

        if($DownloadOnly) {return}        

        $installer = $Session.CreateUpdateInstaller()

        if($installer.IsBusy) {
            Write-Error "Windows Update reports it is busy. Please wait for it to complete before retrying this action."
            return
        }

        if($installer.RebootRequiredBeforeInstallation) {
            Write-Error "Windows Update reports a reboot is required before installing any more updates."
            return
        }

        $installer.AllowSourcePrompts = $false
        $installer.IsForced = $true
        $installer.Updates = $updatesToInstall
        $installer.install() | Out-Null
    }
}

Export-ModuleMember Install-WindowsUpdate