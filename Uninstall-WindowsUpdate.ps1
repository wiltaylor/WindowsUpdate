function Uninstall-WindowsUpdate {
<# 
 .Synopsis
  Uninstalls Windows Updates.

 .Description
  Uninstalls Windows Updates.
 
 .Inputs
  Windows Update objects to uninstall. Retrive these with Get-WindowsUpdate cmdlet.
 
 .Outputs
  Nothing

 .Example
  #Uinstalls all updates in $updates variable.
  $Updates | Install-WindowsUpdate

 .LINK
  about_WindowsUpdateModule

 .LINK
  http://www.win32.io/cmdlet/Uninstall-WindowsUpdate.html
#>
    [CmdletBinding ()]
    param([Parameter(Mandatory=$True,ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$false)]$update)

    BEGIN{
        $UpdateCollection = New-Object -ComObject Microsoft.Update.UpdateColl
    }

    PROCESS {
        $UpdateCollection.Add($update) | Out-Null
    }

    END {
        $Session = New-Object -ComObject Microsoft.Update.Session  

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
        $installer.Updates = $UpdateCollection
        $installer.uninstall() | Out-Null
    }
}

Export-ModuleMember Uninstall-WindowsUpdate