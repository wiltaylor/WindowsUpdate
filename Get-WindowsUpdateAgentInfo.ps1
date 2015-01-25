function Get-WindowsUpdateAgentInfo {
<# 
 .Synopsis
  Retrives information from the Windows Update Agent.

 .Description
  Retrives information from the Windows Update Agent and also retrives WSUS settings.
 
 .Example
  Get-WindowsUpdateAgentInfo

 .LINK
  about_WindowsUpdateModule

 .LINK
  http://www.win32.io/cmdlet/Get-WindowsUpdateAgentInfo.html
#>

    param()

    $agent = New-Object -ComObject Microsoft.Update.AgentInfo
    $autoupdates = New-Object -ComObject Microsoft.Update.AutoUpdate
    $SystemInfo = New-Object -ComObject Microsoft.Update.SystemInfo

    switch($autoupdates.Settings.NotificationLevel) {
        0 { $NotificationLevel =  "Not Configured" }
        1 { $NotificationLevel =  "Disabled" }
        2 { $NotificationLevel =  "Notify Before Download" }
        3 { $NotificationLevel =  "Notify Before Installation" }
        4 { $NotificationLevel =  "Scheduled Installation" }
        default { $NotificationLevel = $autoupdates.Settings.NotificationLevel }
    }

    switch($autoupdates.Settings.ScheduledInstallationDay) {
        0 { $InstallationDay = "Everyday" }
        1 { $InstallationDay = "Sunday" }
        2 { $InstallationDay = "Monday" }
        3 { $InstallationDay = "Tuesday" }
        4 { $InstallationDay = "Wednesday" }
        5 { $InstallationDay = "Thursday" }
        6 { $InstallationDay = "Friday" }
        7 { $InstallationDay = "Saturday" }
        default { $InstallationDay =  $autoupdates.Settings.ScheduledInstallationDay }
    }

    switch($autoupdates.Settings.ScheduledInstallationTime) {
        0  {$ScheduledInstallationTime = "00:00"}
        1  {$ScheduledInstallationTime = "01:00"}
        2  {$ScheduledInstallationTime = "02:00"}
        3  {$ScheduledInstallationTime = "03:00"}
        4  {$ScheduledInstallationTime = "04:00"}
        5  {$ScheduledInstallationTime = "05:00"}
        6  {$ScheduledInstallationTime = "06:00"}
        7  {$ScheduledInstallationTime = "07:00"}
        8  {$ScheduledInstallationTime = "08:00"}
        9  {$ScheduledInstallationTime = "09:00"}
        10 {$ScheduledInstallationTime = "10:00"}
        11 {$ScheduledInstallationTime = "11:00"}
        12 {$ScheduledInstallationTime = "12:00"}
        13 {$ScheduledInstallationTime = "13:00"}
        14 {$ScheduledInstallationTime = "14:00"}
        15 {$ScheduledInstallationTime = "15:00"}
        16 {$ScheduledInstallationTime = "16:00"}
        17 {$ScheduledInstallationTime = "17:00"}
        18 {$ScheduledInstallationTime = "18:00"}
        19 {$ScheduledInstallationTime = "19:00"}
        20 {$ScheduledInstallationTime = "20:00"}
        21 {$ScheduledInstallationTime = "21:00"}
        22 {$ScheduledInstallationTime = "22:00"}
        23 {$ScheduledInstallationTime = "23:00"}
        default { $ScheduledInstallationTime = $autoupdates.Settings.ScheduledInstallationTime}
    }

    if(Test-Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate") {
        $key = Get-Item "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate"
        $WUServer = $key.GetValue("WUServer ", "")
        $WUStatusServer = $key.GetValue("WUStatusServer", "")
        $TargetGroup = $key.GetValue("TargetGroup", "")
        
        if($key.GetValue("AcceptTrustedPublisherCerts", 0) -eq 1) {
            $AcceptTrustedPublisherCerts = $true
        } else {
            $AcceptTrustedPublisherCerts = $false
        }

        if($key.GetValue("DisableWindowsUpdateAccess", 0) -eq 1) {
            $DisableWindowsUpdateAccess = $true
        } else {
            $DisableWindowsUpdateAccess = $false
        }

        if($key.GetValue("TargetGroupEnabled", 0) -eq 1) {
            $TargetGroupEnabled = $true
        } else {
            $TargetGroupEnabled = $false
        }

        
    } else {
        $WUServer = ""
        $WUStatusServer  = ""
        $AcceptTrustedPublisherCerts = $false
        $DisableWindowsUpdateAccess = $false
        $TargetGroup = ""
        $TargetGroupEnabled = $false
    }

    if(Test-Path "HKLM:\Software\Policies\Microsoft\Windows\WindowsUpdate\AU") {
        $key = Get-Item "HKLM:\Software\Policies\Microsoft\Windows\WindowsUpdate\AU"
        
        $DetectionFrequency = $key.GetValue("DetectionFrequency", "Not Set") 
        $RebootRelaunchTimeout = $key.GetValue("RebootRelaunchTimeout", "Not Set") 
        $RebootWarningTimeout = $key.GetValue("RebootWarningTimeout", "Not Set") 
        $RescheduleWaitTime = $key.GetValue("RescheduleWaitTime", "Not Set") 

        if($key.GetValue("UseWUServer", 0) -eq 1) {
            $UseWUServer = $true
        } else {
            $UseWUServer = $false
        }

        if($key.GetValue("DetectionFrequencyEnabled", 0) -eq 1) {
            $DetectionFrequencyEnabled = $true
        } else {
            $DetectionFrequencyEnabled = $false
        }

        if($key.GetValue("NoAutoRebootWithLoggedOnUsers", 0) -eq 1) {
            $NoAutoRebootWithLoggedOnUsers = $true
        } else {
            $NoAutoRebootWithLoggedOnUsers = $false
        }

        if($key.GetValue("RebootRelaunchTimeoutEnabled", 0) -eq 1) {
            $RebootRelaunchTimeoutEnabled = $true
        } else {
            $RebootRelaunchTimeoutEnabled = $false
        }

        if($key.GetValue("RebootWarningTimeoutEnabled", 0) -eq 1) {
            $RebootWarningTimeoutEnabled = $true
        } else {
            $RebootWarningTimeoutEnabled = $false
        }

        if($key.GetValue("RescheduleWaitTimeEnabled", 0) -eq 1) {
            $RescheduleWaitTimeEnabled = $true
        } else {
            $RescheduleWaitTimeEnabled = $false
        }

    } else {
        $UseWUServer = $false
        $DetectionFrequency = "Not Set"
        $DetectionFrequencyEnabled = $false
        $NoAutoRebootWithLoggedOnUsers = $false
        $RebootRelaunchTimeout = "Not Set"
        $RebootRelaunchTimeoutEnabled = $false
        $RebootWarningTimeout = "Not Set"
        $RebootWarningTimeoutEnabled = $false
        $RescheduleWaitTime = "Not Set"
        $RescheduleWaitTimeEnabled = $false
    }

    if(Test-Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer") {
         $key = Get-Item "HKCU:\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
        
        if($key.GetValue("NoWindowsUpdate", 0) -eq 1) {
            $UserWUWebSiteDisabled = $true
        } else {
            $UserWUWebSiteDisabled = $false
        }
    } else {
        $UserWUWebSiteDisabled = $false
    }

    if(Test-Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Policies\WindowsUpdate") {
         $key = Get-Item "HKCU:\Software\Microsoft\Windows\CurrentVersion\Policies\WindowsUpdate"
        
        if($key.GetValue("DisableWindowsUpdateAccess", 0) -eq 1) {
            $WUDisabledByUserPolicy = $true
        } else {
            $WUDisabledByUserPolicy = $false
        }
    } else {
        $WUDisabledByUserPolicy = $false
    }

    if(Test-Path "HKLM:\SYSTEM\Internet Communication Management\Internet Communication") {
         $key = Get-Item "HKLM:\SYSTEM\Internet Communication Management\Internet Communication"
        
        if($key.GetValue("DisableWindowsUpdateAccess", 0) -eq 1) {
            $WUDisabledByComputerPolicy = $true
        } else {
            $WUDisabledByComputerPolicy = $false
        }
    } else {
        $WUDisabledByComputerPolicy = $false
    }
    
    $obj = New-Object PSObject
    $obj | Add-Member -MemberType NoteProperty -Name MajorVersion -Value $agent.GetInfo("ApiMajorVersion")
    $obj | Add-Member -MemberType NoteProperty -Name MinorVersion -Value $agent.GetInfo("ApiMinorVersion")
    $obj | Add-Member -MemberType NoteProperty -Name ProductVersion -Value $agent.GetInfo("ProductVersionString")
    $obj | Add-Member -MemberType NoteProperty -Name AutoUpdatesEnabled -Value $autoupdates.ServiceEnabled
    $obj | Add-Member -MemberType NoteProperty -Name NotificationLevel -Value $NotificationLevel
    $obj | Add-Member -MemberType NoteProperty -Name Required -Value $autoupdates.Settings.Required
    $obj | Add-Member -MemberType NoteProperty -Name ScheduledInstallationDay -Value $InstallationDay
    $obj | Add-Member -MemberType NoteProperty -Name ScheduledInstallationTime -Value $ScheduledInstallationTime
    $obj | Add-Member -MemberType NoteProperty -Name FeaturedUpdatesEnabled -Value $autoupdates.Settings.FeaturedUpdatesEnabled
    $obj | Add-Member -MemberType NoteProperty -Name NonAdministratorsElevated -Value $autoupdates.Settings.NonAdministratorsElevated
    $obj | Add-Member -MemberType NoteProperty -Name IncludeRecommendedUpdates -Value $autoupdates.Settings.IncludeRecommendedUpdates
    $obj | Add-Member -MemberType NoteProperty -Name LastInstallationSuccessDate -Value $autoupdates.Results.LastInstallationSuccessDate
    $obj | Add-Member -MemberType NoteProperty -Name LastSearchSuccessDate -Value $autoupdates.Results.LastSearchSuccessDate
    $obj | Add-Member -MemberType NoteProperty -Name RebootRequired -Value $SystemInfo.RebootRequired
    $obj | Add-Member -MemberType NoteProperty -Name OEMHardwareSupportLink -Value $SystemInfo.OemHardwareSupportLink
    $Obj | Add-Member -MemberType NoteProperty -Name WUSServer -Value $WUServer
    $Obj | Add-Member -MemberType NoteProperty -Name WUStatusServer -Value $WUStatusServer
    $Obj | Add-Member -MemberType NoteProperty -Name UseWUServer -Value $UseWUServer
    $Obj | Add-Member -MemberType NoteProperty -Name AcceptTrustedPublisherCerts -Value $AcceptTrustedPublisherCerts
    $Obj | Add-Member -MemberType NoteProperty -Name DisableWindowsUpdateAccess -Value $DisableWindowsUpdateAccess
    $Obj | Add-Member -MemberType NoteProperty -Name TargetGroup -Value $TargetGroup
    $Obj | Add-Member -MemberType NoteProperty -Name TargetGroupEnabled -Value $TargetGroupEnabled
    $Obj | Add-Member -MemberType NoteProperty -Name DetectionFrequency -Value $DetectionFrequency
    $Obj | Add-Member -MemberType NoteProperty -Name NoAutoRebootWithLoggedOnUsers -Value $NoAutoRebootWithLoggedOnUsers
    $Obj | Add-Member -MemberType NoteProperty -Name RebootRelaunchTimeout -Value $RebootRelaunchTimeout
    $Obj | Add-Member -MemberType NoteProperty -Name RebootRelaunchTimeoutEnabled -Value $RebootRelaunchTimeoutEnabled
    $Obj | Add-Member -MemberType NoteProperty -Name RebootWarningTimeout -Value $RebootWarningTimeout
    $Obj | Add-Member -MemberType NoteProperty -Name RebootWarningTimeoutEnabled -Value $RebootWarningTimeoutEnabled
    $Obj | Add-Member -MemberType NoteProperty -Name RescheduleWaitTime -Value $RescheduleWaitTime
    $Obj | Add-Member -MemberType NoteProperty -Name RescheduleWaitTimeEnabled -Value $RescheduleWaitTimeEnabled
    $Obj | Add-Member -MemberType NoteProperty -Name UserWUWebSiteDisabled -Value $UserWUWebSiteDisabled
    $Obj | Add-Member -MemberType NoteProperty -Name WUDisabledByUserPolicy -Value $WUDisabledByUserPolicy
    $Obj | Add-Member -MemberType NoteProperty -Name WUDisabledByComputerPolicy -Value $WUDisabledByComputerPolicy
    
    #First Parma of checkPermisions 1 means current user. Second Param is AutomaticUpdatesPermissionType enumeration
    $obj | Add-Member -MemberType NoteProperty -Name CanSetNotificationLevel -Value $autoupdates.Settings.CheckPermission(1, 1)
    $obj | Add-Member -MemberType NoteProperty -Name CanDisableAutomaticUpdates -Value $autoupdates.Settings.CheckPermission(1, 2)
    $obj | Add-Member -MemberType NoteProperty -Name CanSetIncludeRecommendedUpdates -Value $autoupdates.Settings.CheckPermission(1, 3)
    #$obj | Add-Member -MemberType NoteProperty -Name CanSetFeaturedUpdatesEnabled -Value $autoupdates.Settings.CheckPermission(1, 4) - Generates com errors.
    $obj | Add-Member -MemberType NoteProperty -Name CanSetNonAdministratorsElevated -Value $autoupdates.Settings.CheckPermission(1, 5)


    $obj
}

Export-ModuleMember Get-WindowsUpdateAgentInfo
