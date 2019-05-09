# Set path and name variables
$ProfileRegPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
$RegLoadName = 1
$HiveFile = "NTUSER.DAT"
$RegLoadPath = "registry::HKEY_USERS"

$RunRegPath = "SOFTWARE\Microsoft\Windows\Currentversion\Run"
$PeopleRegPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
$NotificationsRegPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Notifications\Settings"
$PushNotifRegPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\PushNotifications\Backup"

$ODRegProp = "OneDriveSetup"

$PeopleRegKey = "People"
$PeopleBandRegProp = "PeopleBand"

$LockScreenRegProp = "NOC_GLOBAL_SETTING_ALLOW_TOASTS_ABOVE_LOCK"
$LockVOIPRegProp = "NOC_GLOBAL_SETTING_ALLOW_CRITICAL_TOASTS_ABOVE_LOCK"
$DupScreenRegProp = "NOC_GLOBAL_SETTING_SUPRESS_TOASTS_WHILE_DUPLICATING"

$SugRegKey = "Windows.SystemToast.Suggested"

$SugAppTypeRegProp = "appType"
$SugAppTypeRegValue = "app:system"

$SugWnsldRegProp = "wnsld"
$SugWnsldRegValue = "System"

$SugSettingRegProp = "Setting"
$SugSettingRegValue = "c:toast,c:ringing,c:storage:toast,s:tickle,s:toast,s:audio,s:badge,s:lock:badge,s:banner,s:listenerEnabled,s:lock:tile,s:tile,s:lock:toast,s:voip" 

# Get profile path info for all user profiles with SIDs beginning with S-1-5-21 (normal domain and local users)
$ProfileKeys = Get-ChildItem -Path $ProfileRegPath
$Profiles = $ProfileKeys | ForEach-Object {Get-ItemProperty -Path $_.PSPath | Where-Object {$_ -match "S-1-5-21-*"}}

# For each profile, mount the NTUSER.DAT hive, delete the OneDriveSetup Run property if found, and unmount the hive
foreach ($Profile in $Profiles) {
    
    $ImagePath = $Profile.ProfileImagePath

    reg load HKU\$($RegLoadName.ToString()) $ImagePath\$HiveFile | Out-Null

    # Checks for the existence of the OneDriveSetup property, whether it has a value or not
    if ((Get-ItemProperty -Path $RegLoadPath\$($RegLoadName.ToString())\$RunRegPath).PSObject.Properties.Name -contains $ODRegProp) {
    
        Remove-ItemProperty -Path $RegLoadPath\$($RegLoadName.ToString())\$RunRegPath -Name $ODRegProp
    }

    # Add People key and PeopleBand property to disable the taskbar people button
    New-Item -Path $RegLoadPath\$($RegLoadName.ToString())\$PeopleRegPath -Name $PeopleRegKey -Force
    New-ItemProperty -Path $RegLoadPath\$($RegLoadName.ToString())\$PeopleRegPath\$PeopleRegKey -Name $PeopleBandRegProp -PropertyType DWORD -Value 0 -Force

    # Configure notification settings to hide on lock screen and when duplicating the screen
    New-ItemProperty -Path $RegLoadPath\$($RegLoadName.ToString())\$NotificationsRegPath -Name $LockScreenRegProp -PropertyType DWORD -Value 0 -Force
    New-ItemProperty -Path $RegLoadPath\$($RegLoadName.ToString())\$NotificationsRegPath -Name $LockVOIPRegProp -PropertyType DWORD -Value 0 -Force
    New-ItemProperty -Path $RegLoadPath\$($RegLoadName.ToString())\$NotificationsRegPath -Name $DupScreenRegProp -PropertyType DWORD -Value 1 -Force

    # Turn off notifications from the "Suggested" sender
    New-Item -Path $RegLoadPath\$($RegLoadName.ToString())\$NotificationsRegPath -Name $SugRegKey -Force
    New-ItemProperty -Path $RegLoadPath\$($RegLoadName.ToString())\$NotificationsRegPath\$SugRegKey -Name Enabled -PropertyType DWORD -Value 0 -Force

    # Configure the "Suggested" sender to show in the "Get notifications from these senders" list
    New-Item -Path $RegLoadPath\$($RegLoadName.ToString())\$PushNotifRegPath -Name $SugRegKey -Force
    New-ItemProperty -Path $RegLoadPath\$($RegLoadName.ToString())\$PushNotifRegPath\$SugRegKey -Name $SugAppTypeRegProp -PropertyType String -Value $SugAppTypeRegValue -Force
    New-ItemProperty -Path $RegLoadPath\$($RegLoadName.ToString())\$PushNotifRegPath\$SugRegKey -Name $SugWnsldRegProp -PropertyType String -Value $SugWnsldRegValue -Force
    New-ItemProperty -Path $RegLoadPath\$($RegLoadName.ToString())\$PushNotifRegPath\$SugRegKey -Name $SugSettingRegProp -PropertyType String -Value $SugSettingRegValue -Force

    # Run garbage collection to free up any handles that may prevent the hive from successfully unloading
    [gc]::Collect()
    reg unload HKU\$($RegLoadName.ToString()) | Out-Null

    # Increment the number used for the load key name so it is unique for each iteration
    $RegLoadName++
}