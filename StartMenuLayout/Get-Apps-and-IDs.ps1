#this is the Logoff script
Param(
    # Working Folder for files. Should be the same location you define Group Policy to look for the 'layout.xml, This will try and grab the value from the registry'
    [Parameter(Mandatory = $false)]
    $layoutPath = (Split-Path -path (Get-ItemPropertyValue -Path "HKCU:\Software\Policies\Microsoft\Windows\Explorer\" -Name StartLayoutFile)) + "\"
)

#layout path here is commented out as the information is coming from the windows registry for the layout.xml storage location. 
#this works ok when using a layout that is stored localy on a computer. If its remote this will collide with other pc's
#to resolve any issues set the parameter on input or set hard coded here:
#$layoutPath = "C:\ProgramData\Startmenu" 
$startAppsJson = "AvailableStartApps.json"

$pubIDs = Get-AppxPackage | Select-Object -ExpandProperty PublisherID | Sort-Object | Get-Unique
$availableStartApps = Get-StartApps
 
$availableStartApps | Add-Member -NotePropertyName TileType -NotePropertyValue ""
$availableStartApps | Add-Member -NotePropertyName AppIDType -NotePropertyValue ""

foreach ($app in $availableStartApps) {
    #finding if an item contains a publisher ID in APPID or not. Those with PublisherID's are a Tile Type application (ie probably from Microsoft Store)
    if ($app.AppID | Select-String $pubIDs -Quiet) {
        $app.TileType = "Tile"
        $app.AppIDType = "AppUserModelID"
    }
    else {
        $app.TileType = "DesktopApplicationTile"
        $app.AppIDType = "DesktopApplicationID"
    }
}

$availableStartApps | ConvertTo-Json | Out-File -FilePath $layoutPath\$startAppsJson