$LayoutPath = "C:\StartMenu"
$StartAppsFile = "StartApps.csv"
$PublisherIDsFile = "PublisherIDs.txt"

Get-StartApps | Export-Csv -Path $LayoutPath\$StartAppsFile

Get-AppxPackage | Select-Object -ExpandProperty PublisherID | Sort-Object | Get-Unique | Out-File -FilePath $LayoutPath\$PublisherIDsFile
