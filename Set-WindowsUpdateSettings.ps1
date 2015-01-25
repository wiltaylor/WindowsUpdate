function Set-WindowsUpdateSettings {
<# 
 .Synopsis
  Changes a windows update setting.

 .Description
  Changes a windows update setting.

 .Parameter EnableAutoUpdate
  This switch enables automatic updates.

 .Parameter DisableAutoUpdate
  This switch disables automatic updates.

 .Parameter EnableFeaturedUpdates
  This switch enables Featured Updates from being installed by automatic updates.

 .Parameter DisableFeaturedUpdates
  This switch disables Featured Updates from being installed by automatic updates.

 .Parameter EnableRecommendedUpdates
  This switch enables Recommended Updates from being installed by automatic updates.

 .Parameter DisableRecommendedUpdates
  This switch disables Recommended Updates from being installed by automatic updates.

 .Parameter EnableNonAdministrators
  This switch allows non administrative users to run windows automatic updates.

 .Parameter DisableNonAdministrators
  This switch prevents non administrative users to run windows automatic updates.

 .Parameter NotificationLevel
  This sets the level Automatic updates runs at. Available options are the following:
    Not Configured
    Disabled
    Notify Before Download
    Notify Before Installation
    Scheduled Installation
 
 .Parameter InstallationDay
  This is the day of the week which Automatic updates will run. Available options are the following:
    Everyday
    Sunday
    Monday
    Tuesday
    Wednesday
    Thursday
    Friday
    Saturday

 .Parameter InstallTime
  This is the time during the day that patches will install. Enter this in the following format HH:MM.
  For example 12:00. You also can only enter exactly on the hour. i.e. 1:00 is valid where as 1:09 is not.

 .Example
  Set-WindowsUpdateSettings -EnableAutoUpdate -EnableFeaturedUpdates -DisableRecommendedUpdates

 .Example
  Set-WindowsUpdateSettings -DisableNonAdministrators -NotificationLevel Everyday -InstallTime 03:00

 .LINK
  about_WindowsUpdateModule

 .LINK
  http://www.win32.io/cmdlet/Set-WindowsUpdateSettings.html
#>
    param([switch]$EnableAutoUpdate, [switch]$DisableAutoUpdate, [switch]$EnableFeaturedUpdates, [switch]$DisableFeaturedUpdates, 
    [switch]$EnableRecommendedUpdates, [switch]$DisableRecommendedUpdates, [switch]$EnableNonAdministrators, [switch]$DisableNonAdministrators,
    [ValidateSet("","Not Configured", "Disabled", "Notify Before Download", "Notify Before Installation", "Scheduled Installation")]
    [string]$NotificationLevel="", 
    [ValidateSet("","Everyday", "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday")]
    [string]$InstallationDay="",
    [ValidateSet("","00:00", "01:00", "02:00", "03:00", "04:00", "05:00", "06:00", "07:00", "08:00", "09:00", "10:00", "11:00", "12:00", "13:00", "14:00", 
    "15:00", "16:00", "17:00", "18:00", "19:00", "20:00", "21:00", "22:00", "23:00")]
    [string]$InstallTime="")

    $autoupdates = New-Object -ComObject Microsoft.Update.AutoUpdate

    if($autoupdates.Settings.ReadOnly) {
        Write-Error "Windows Update Reports settings are read only! Check that you have enough security rights to make changes."
        return
    }

    if($EnableAutoUpdate) {
        $autoupdates.EnableService() | Out-Null
    }

    if($DisableAutoUpdate) {
        $NotificationLevel = "Disabled"
    }

    if($EnableFeaturedUpdates) {
        $autoupdates.Settings.FeaturedUpdatesEnabled = $true
    }

    if($DisableFeaturedUpdates) {
        $autoupdates.Settings.FeaturedUpdatesEnabled = $false
    }

    if($EnableNonAdministrators) {
        $autoupdates.Settings.NonAdministratorsElevated = $true
    }

    if($DisableNonAdministrators) {
        $autoupdates.Settings.NonAdministratorsElevated = $false
    }

    if($EnableRecommendedUpdates) {
        $autoupdates.Settings.IncludeRecommendedUpdates = $true
    }

    if($DisableRecommendedUpdates) {
        $autoupdates.Settings.IncludeRecommendedUpdates = $false
    }
    
    if($InstallationDay -ne "") {
        switch($InstallationDay) {
            "Everyday"  { $day = 0 }
            "Sunday"    { $day = 1 }
            "Monday"    { $day = 2 }
            "Tuesday"   { $day = 3 }
            "Wednesday" { $day = 4 }
            "Thursday"  { $day = 5 }
            "Friday"    { $day = 6 }
            "Saturday"  { $day = 7 }
        }

        $autoupdates.Settings.ScheduledInstallationDay = $day
        
    }

    if($InstallTime -ne "") {
        switch($InstallTime) {
            "00:00" {$ScheduledInstallationTime = 0}
            "01:00" {$ScheduledInstallationTime = 1}
            "02:00" {$ScheduledInstallationTime = 2}
            "03:00" {$ScheduledInstallationTime = 3}
            "04:00" {$ScheduledInstallationTime = 4}
            "05:00" {$ScheduledInstallationTime = 5}
            "06:00" {$ScheduledInstallationTime = 6}
            "07:00" {$ScheduledInstallationTime = 7}
            "08:00" {$ScheduledInstallationTime = 8}
            "09:00" {$ScheduledInstallationTime = 9}
            "10:00" {$ScheduledInstallationTime = 10}
            "11:00" {$ScheduledInstallationTime = 11}
            "12:00" {$ScheduledInstallationTime = 12}
            "13:00" {$ScheduledInstallationTime = 13}
            "14:00" {$ScheduledInstallationTime = 14}
            "15:00" {$ScheduledInstallationTime = 15}
            "16:00" {$ScheduledInstallationTime = 16}
            "17:00" {$ScheduledInstallationTime = 17}
            "18:00" {$ScheduledInstallationTime = 18}
            "19:00" {$ScheduledInstallationTime = 19}
            "20:00" {$ScheduledInstallationTime = 20}
            "21:00" {$ScheduledInstallationTime = 21}
            "22:00" {$ScheduledInstallationTime = 22}
            "23:00" {$ScheduledInstallationTime = 23}
        }

        $autoupdates.Settings.ScheduledInstallationTime = $ScheduledInstallationTime
    }

    if($NotificationLevel -ne "") {
        switch($NotificationLevel) {
            "Not Configured"             {$level = 0}
            "Disabled"                   {$level = 1}
            "Notify Before Download"     {$level = 2}
            "Notify Before Installation" {$level = 3}
            "Scheduled Installation"     {$level = 4}
        }

        $autoupdates.Settings.NotificationLevel = $level
    }

    $autoupdates.Settings.Save() | Out-Null

}

Export-ModuleMember Set-WindowsUpdateSettings