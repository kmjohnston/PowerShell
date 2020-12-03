#this is the Logon script.
Param(
    # Working Folder for files. Should be the same location you define Group Policy to look for the "layout.xml", This will try and grab the value from the registry'
    [Parameter(Mandatory = $false)]
    $layoutPath = (Split-Path -path (Get-ItemPropertyValue -Path "HKCU:\Software\Policies\Microsoft\Windows\Explorer\" -Name StartLayoutFile)) + "\"
)
#set global Error actions
$ErrorActionPreference = "Stop"
$ErrorView = "CategoryView"

# Set file and folder names
#the layout path does not need to be set here unless the layout path gpo is not used or if the xml path is stored on a network drive used by multiple users(why wold you use this script then)
#$layoutPath = "C:\ProgramData\StartMenu\"

if (!(Test-Path -path $layoutPath)) {
    Write-Error "No usable layout path found"
    Write-Error $error[0].exception 
}
$layoutXMLFile = $layoutPath + "Layout.xml"
$tempLayoutXMLFile = $layoutPath + "Layout_temp.xml"

$availableStartAppsJson = $layoutPath + "AvailableStartApps.json"
$menuListJson = $layoutPath + "Menu.json"
    
# Load the valid start Apps list
try {
    $availablestartApps = Get-Content -path $availableStartAppsJson | ConvertFrom-Json
}
catch {
    Write-Warning "Unable to load StartApps from $availableStartAppsJson"
    Write-Error $_
}
# Load the menu groups
try {
    $menuGroups = Get-Content $menuListJson | ConvertFrom-Json
}
catch {
    Write-Warning "!!!Unable to load Menu Groups from: $menuListJson"
    Write-Error $_
}

function WriteShellXML {
    
    # Write the required XML elements for the Start layout file
    $writer.WriteStartElement("LayoutModificationTemplate", "http://schemas.microsoft.com/Start/2014/LayoutModification")  
    $writer.WriteAttributeString("Version", "1")

    $writer.WriteStartElement("LayoutOptions")
    $writer.WriteAttributeString("StartTileGroupCellWidth", "6")
    $writer.WriteEndElement()

    $writer.WriteStartElement("DefaultLayoutOverride")
    $writer.WriteAttributeString("LayoutCustomizationRestrictionType", "OnlySpecifiedGroups")

    $writer.WriteStartElement("StartLayoutCollection")

    $writer.WriteStartElement("defaultlayout", "StartLayout", "http://schemas.microsoft.com/Start/2014/FullDefaultLayout")
    $writer.WriteAttributeString("GroupCellWidth", "6")
}

function WriteTileXML ($group) {
    # Write the Start Group XML Element
    $writer.WriteStartElement("start", "Group", "http://schemas.microsoft.com/Start/2014/StartLayout")
    $writer.WriteAttributeString("Name", $group.name)
    
    #this looses the order of the apps from menu.json/$group.members
    $confirmedStartApps = $availablestartApps | Where-Object { $_.Name -in $group.members }

    $sortedConfirmedStartApps = @()
    foreach ($member in $group.members) {
        $sortedConfirmedStartApps += $confirmedStartApps | Where-Object { $_.name -eq $member }
    }
    
    # Loop through the group apps list and write corresponding XML elements
    foreach ($confirmedApp in $sortedConfirmedStartApps) {
        #$confirmedApp | select AppID, Name, TileType
        $index = $sortedConfirmedStartApps.IndexOf($confirmedApp)

        $row = [math]::Truncate($index / 3) * 2
        $column = ($index % 3) * 2
        Write-Output "row: $row   column: $column     $($confirmedApp.Name)"

        #Write XML elements and attributes for the tile
        $writer.WriteStartElement("start", $confirmedApp.TileType, "http://schemas.microsoft.com/Start/2014/StartLayout")
        $writer.WriteAttributeString("Size", "2x2")
        $writer.WriteAttributeString("Column", $column)
        $writer.WriteAttributeString("Row", $row)
        $writer.WriteAttributeString($confirmedApp.AppIDType, $confirmedApp.AppID)
        $writer.WriteEndElement()
    }
    $writer.WriteEndElement()            
}

# Create a new XML writer settings object and configure settings 
$settings = New-Object system.Xml.XmlWriterSettings 
$settings.Indent = $true 
$settings.OmitXmlDeclaration = $true

# Create a new XML writer 
Try {
    $writer = [system.xml.XmlWriter]::Create($tempLayoutXMLFile, $settings)

    # Call function to write XML shell
    WriteShellXML
    $menuGroups.menuGroup | ForEach-Object {
        WriteTileXML -group $_
    }
    # Flush the XML writer and close the file     
    $writer.Flush() 
    $writer.Close()
}
Catch {
    Write-Warning "Atempted to write the xml but ran into an issue:"
    Write-Error $_
}

if (Test-Path -path $layoutXMLFile) {
    #compares the new and old XML 
    $FilesDifferent = Compare-Object $(Get-Content -LiteralPath $layoutXMLFile) $(Get-Content -LiteralPath $tempLayoutXMLFile )
    Write-Output $FilesDifferent
}
else {
    $FilesDifferent = $true
}

if ($FilesDifferent) {
    Write-Verbose "Start Menu Layout Updated"
    Copy-Item $tempLayoutXMLFile $layoutXMLFile -verbose -Force
    if (!$?) { throw $error[0].exception }
}
else {
    Write-Verbose "Start Menu Layout Not Updated"
}