<#
.Synopsis
   Use this cmdlet to discover and download Dell BIOS/Driver updates
.DESCRIPTION
   This cmdlet requests website data from Dell's drivers & downloads simplified interface site (http://downloads.dell.com/published/Pages/)
   and enables you to discover and download related files for a specific model or list of models
.EXAMPLE
   Get a list of all available BIOS downloads for a specific model and present the results in a gridview:

   Get-DellDownloads -ModelName "Latitude E7470" -Category BIOS -Type BIOS 
.EXAMPLE
   Get a list of all available video driver downloads for a specific model and present the results in a gridview, including the description column. Download the selected files from that list to the location specified:

   Get-DellDownloads -ModelName "Latitude E7240 Ultrabook" -Category Video -Type Driver -IncludeDescription -DownloadLocation C:\Temp
.EXAMPLE
   For a provided text file containing a list of multiple models, find possible categories/types and present the results in a griview. Based on the selection from that list, get a list of the most recent downloads of that category/type for each model, and present a gridview list of the results. Download the selected files from that list to the location specified:

   Get-DellDownloads -ModelList C:\Temp\ModelList.txt -DownloadLocation C:\Temp
.EXAMPLE
   For a provided text file containing a list of multiple models, get a list of the most recent BIOS downloads for each model and present the results in a gridview:

   Get-DellDownloads -ModelList C:\Temp\ModelList.txt -Category BIOS -Type BIOS
.NOTES
   Author: Kevin Johnston
   Date:   December 10, 2018
#>
function Get-DellDownloads
{
   [CmdletBinding()]
   Param
    (
        [parameter(Mandatory=$true,ParameterSetName="Name")][String]$ModelName,
        [parameter(Mandatory=$true,ParameterSetName="List")][String]$ModelList,
        [parameter(Mandatory=$false)]
        [ValidateSet("Application","Audio","Backup and Recovery","BIOS","Chipset","CMDSK","Communications",`
                     "Dell Data Protection","Docks/Stands","Drivers for OS Deployment","Firmware","IDM","Input",`
                     "Miscellaneous Utilities","Network","Removeable Storage","SAS Drive","SAS RAID","Security Encryption",`
                     "Serial ATA","Storage","Storage Controller","Systems Management","System Utilities","Video")]
        [String]$Category,
        [parameter(Mandatory=$false)]
        [ValidateSet("Application","BIOS","Diagnostics Utility","Driver","Firmware","HTML","ISV Driver","Utility")]
        [String]$Type,
        [switch]$IncludeDescription,
        [parameter(Mandatory=$false)][String]$DownloadLocation
    )

    # set models variable to single model or list depending on parameter used
    if ($ModelName) {$models = $ModelName}
    if ($ModelList) {$models = Get-Content -Path $ModelList}

    if ($Category -and $Type)
    {
        # create array of known SectionID, Category, and Type info 
        $knownSections = @"
"Drivers-Category.AP-Type.APP","Application","Application"
"Drivers-Category.AS-Type.FRMW","SAS Drive","Firmware"
"Drivers-Category.AU-Type.DRVR","Audio","Driver"
"Drivers-Category.BI-Type.BIOS","BIOS","BIOS"
"Drivers-Category.BR-Type.APP","Backup and Recovery","Application"
"Drivers-Category.CM-Type.APP","Communications","Application"
"Drivers-Category.CM-Type.DRVR","Communications","Driver"
"Drivers-Category.CM-Type.UTIL","Communications","Utility"
"Drivers-Category.CS-Type.APP","Chipset","Application"
"Drivers-Category.CS-Type.DRVR","Chipset","Driver"
"Drivers-Category.CS-Type.FRMW","Chipset","Firmware"
"Drivers-Category.DD-Type.APP","Drivers for OS Deployment","Application"
"Drivers-Category.DD-Type.DRVR","Drivers for OS Deployment","Driver"
"Drivers-Category.DK-Type.DRVR","Docks/Stands","Driver"
"Drivers-Category.DK-Type.FRMW","Docks/Stands","Firmware"
"Drivers-Category.DK-Type.UTIL","Docks/Stands","Utility"
"Drivers-Category.DP-Type.APP","Dell Data Protection","Application"
"Drivers-Category.DP-Type.DRVR","Dell Data Protection","Driver"
"Drivers-Category.FW-Type.FRMW","Firmware","Firmware"
"Drivers-Category.IDM-Type.APP","IDM","Application"
"Drivers-Category.IN-Type.APP","Input","Application"
"Drivers-Category.IN-Type.DRVR","Input","Driver"
"Drivers-Category.MU-Type.UTIL","Miscellaneous Utilities","Utility"
"Drivers-Category.NI-Type.APP","Network","Application"
"Drivers-Category.NI-Type.DIAG","Network","Diagnostics Utility"
"Drivers-Category.NI-Type.DRVR","Network","Driver"
"Drivers-Category.NI-Type.FRMW","Network","Firmware"
"Drivers-Category.NI-Type.HTML","Network","HTML"
"Drivers-Category.NI-Type.UTIL","Network","Utility"
"Drivers-Category.RS-Type.FRMW","Removable Storage","Firmware"
"Drivers-Category.SA-Type.DRVR","Serial ATA","Driver"
"Drivers-Category.SA-Type.FRMW","Serial ATA","Firmware"
"Drivers-Category.SA-Type.UTIL","Serial ATA","Utility"
"Drivers-Category.SF-Type.APP","SAS RAID","Application"
"Drivers-Category.SF-Type.FRMW","SAS RAID","Firmware"
"Drivers-Category.SG-Type.APP","Storage Controller","Application"
"Drivers-Category.SG-Type.DRVR","Storage Controller","Driver"
"Drivers-Category.SK-Type.APP","CMDSK","Application"
"Drivers-Category.SM-Type.APP","Systems Management","Application"
"Drivers-Category.SM-Type.DRVR","Systems Management","Driver"
"Drivers-Category.SM-Type.UTIL","Systems Management","Utility"
"Drivers-Category.SP-Type.APP","Security Encryption","Application"
"Drivers-Category.SP-Type.DRVR","Security Encryption","Driver"
"Drivers-Category.ST-Type.DRVR","Storage","Driver"
"Drivers-Category.SY-Type.DRVR","Security","Driver"
"Drivers-Category.SY-Type.FRMW","Security","Firmware"
"Drivers-Category.UT-Type.APP","System Utilities","Application"
"Drivers-Category.UT-Type.DRVR","System Utilities","Driver"
"Drivers-Category.UT-Type.UTIL","System Utilities","Utility"
"Drivers-Category.VI-Type.APP","Video","Application"
"Drivers-Category.VI-Type.DRVR","Video","Driver"
"Drivers-Category.VI-Type.FRMW","Video","Firmware"
"Drivers-Category.VI-Type.ISV","Video","ISV Driver"
"Drivers-Category.VI-Type.UTIL","Video","Utility"
"@ -split [System.Environment]::NewLine | ConvertFrom-Csv -Header SectionID,Category,Type
        
        # if the Category and Type parameters are used, set the corresponding SectionID variable
        $section = $knownSections | Where-Object {($_.Category -eq $Category) -and ($_.Type -eq $Type)}
        $sectionID = $section.SectionID
    }

    # set URI variables
    $baseURI = "http://downloads.dell.com/published/Pages/"
    $indexURI = $baseURI + "index.html"

    # request the download index webpage
    $dlIndex = Invoke-WebRequest -Uri $indexURI

    # get all links from the webpage
    $indexLinks = $dlIndex.Links
    
    # if the Category and Type paremeters were not used to define a SectionID,
    # request all possible values from the model(s) specified
    if (-not($sectionID))
    {
       # initialize array to store final unique category/type/sectionID info
       $categoryResults = @()

        foreach ($model in $models)
        {
            # set the link variable for the specific model webpage
            $modelLink = $indexLinks | Where-Object {$_.innerHTML -eq $model}

            # set the URI variable for the specific model webpage
            $modelURI = $baseURI + $modelLink.href

            # request the specific model webpage
            $modelIndex = Invoke-WebRequest -Uri $modelURI
            
            Write-Output "Requesting available categories and types for model $model"

            # get webpage elements for the model sections
            $sectionIndex = $modelIndex.ParsedHTML.getElementsByTagName('DIV')
            
            # get all of the SectionIDs with '-Type.' in the name
            $typeIDs = $sectionIndex | Where-Object {$_.id -like '*-Type.*'}

            # get webpage elements for the model headings
            $headingIndex = $modelIndex.ParsedHTML.getElementsByTagName('H5')

            # initialize an empty array to store category/type/sectionID info
            $categories = @()

            # initialize a counter for the SectionID types array
            $idCounter = 0

            # loop through the model headings
            for ($headingCounter = 0; $headingCounter -lt ($headingIndex | Measure-Object).Count; $headingCounter++)
            { 
                if ($headingIndex[$headingCounter].innerText -like 'Category:*')
                {
                    # if the heading is a category, set the category text variable
                    $charIndex = ($headingIndex[$headingCounter].innerText).IndexOf(':')
                    $categoryText = ($headingIndex[$headingCounter].innerText).Substring($charIndex + 2)
                }

                if ($headingIndex[$headingCounter].innerText -like 'Type:*')
                {
                    # if the heading is a type, set the type text variable
                    $charIndex = ($headingIndex[$headingCounter].innerText).IndexOf(':')
                    $typeText = ($headingIndex[$headingCounter].innerText).Substring($charIndex + 2)

                    # set the corresponding sectionID text variable
                    $sectionIDtext = $typeIDs[$idCounter].id

                    # add category, type, and sectionID 
                    $categories += New-Object psobject -Property @{Category=$categoryText;Type=$typeText;SectionID=$sectionIDtext}

                    # increment the sectionID types counter
                    $idCounter++
                }
            }
            
            # loop through the category/type/sectionID array 
            foreach ($item in $categories)
            {
                # if the category results array does not already contain an entry with the element's section ID, add to the array 
                if ($categoryResults.sectionID -notcontains $item.sectionID)
                {
                    $categoryResults += $item
                }
            }
        }

        # display the results in a grid view and set the sectionID variable to the user's selection
        $sectionID = $categoryResults | Select-Object -Property Category,Type,SectionID | Sort-Object -Property Category | Out-GridView -PassThru | Select-Object -ExpandProperty SectionID 
    }

    # initialize an empty array to store model results
    $modelResults = @()

    foreach ($model in $models)
    {
        # set the link variable for the specific model webpage
        $modelLink = $indexLinks | Where-Object {$_.innerHTML -eq $model}

        # set the URI variable for the specific model webpage
        $modelURI = $baseURI + $modelLink.href

        # request the specific model webpage
        $modelIndex = Invoke-WebRequest -Uri $modelURI

        Write-Output "Requesting available downloads for model $model"

        # get webpage elements for the desired section ID
        $sectionIndex = $modelIndex.ParsedHtml.getElementsByTagName('DIV') | Where-Object {$_.id -eq $sectionID}

        # get webpage elements for the section rows
        $sectionRows = $sectionIndex.getElementsByTagName('TR')

        # initialize an empty array to store section results
        $sectionResults = @()

        # loop through each section row (skipping the first which only contains known header values)
        for ($secCounter = 1; $secCounter -lt ($sectionRows | Measure-Object).Count; $secCounter++)
        { 
            # get webpage elements for the row cells
            $sectionCells = $sectionRows[$secCounter].getElementsByTagName('TD')

            # loop through each row cell
            for ($cellCounter = 0; $cellCounter -lt ($sectionCells | Measure-Object).Count; $cellCounter++)
            { 
                # set Download cell value(s)
                if ($cellCounter -eq 5)
                {
                    # get hyperlink webpage elements for the download cell
                    $cellLinks = $sectionCells[$cellCounter].getElementsByTagName('A')
                
                    # get the download links and change them to https (seems to work better for actual downloading)
                    $dlLinks = ($cellLinks | Select-Object -ExpandProperty href) -replace 'http://','https://'
                
                    if ($dlLinks.Count -gt 1)
                    {
                        # for cells with multiple links, convert array to single string with newlines.
                        # this allows the final results to display like the other cells
                        $dlLinks = ($dlLinks -join [Environment]::NewLine | Out-String).TrimEnd()
                    }
                }
                else
                {
                    # set other cell values
                    switch ($cellCounter)
                    {
                        '0' {$DescriptionText = $sectionCells[$cellCounter].innerText}
                        '1' {$Importance = $sectionCells[$cellCounter].innerText}
                        '2' {$Version = $sectionCells[$cellCounter].innerText}
                        '3' {$Released = ($sectionCells[$cellCounter].innerText | Get-Date)}
                        '4' {$SupportedOS = $sectionCells[$cellCounter].innerText}
                    }
                }
            }

            # add cell values for each row to the section results array
            $sectionResults += New-Object psobject -Property @{Description=$DescriptionText;
                                                               Importance=$Importance;
                                                               Version=$Version;
                                                               Released=$Released;
                                                               SupportedOS=$SupportedOS;
                                                               Download=$dlLinks}
        }
        
        # if more than one model, get only the most recent release(s) for each model
        if ($models.Count -gt 1)
        {
             # set variable for the latest date found in the section results array
            $latestDate = ($sectionResults.Released | Measure-Object -Maximum).Maximum

            # set variable for the latest release(s) found that match(es) the latest date
            $latestRelease = $sectionResults | Where-Object {$_.Released -eq $latestDate}

            foreach ($release in $latestRelease)
            {   
                # add the latest release row(s) to the model results array
                $modelResults += New-Object psobject -Property @{Model=$model;
                                                                 Description=$release.Description;
                                                                 Importance=$release.Importance;
                                                                 Version=$release.Version;
                                                                 Released=$release.Released;
                                                                 SupportedOS=$release.SupportedOS;
                                                                 Download=$release.Download}
            } 
        }
        # else get all releases for the single model
        else
        {
            foreach ($sectionItem in $sectionResults)
            {   
                # add the release rows to the model results array
                $modelResults += New-Object psobject -Property @{Model=$model;
                                                                 Description=$sectionItem.Description;
                                                                 Importance=$sectionItem.Importance;
                                                                 Version=$sectionItem.Version;
                                                                 Released=$sectionItem.Released;
                                                                 SupportedOS=$sectionItem.SupportedOS;
                                                                 Download=$sectionItem.Download}
            }
        }
    }
    
    # sort results by date
    $sortedResults = $modelResults | Sort-Object -Property Released -Descending

    # change the Released datetimes to short date strings so the unnecessary time part doesn't display
    $sortedResults | ForEach-Object {$_.Released = $_.Released.ToShortDateString()}

    # define desired properties to display
    if ($IncludeDescription) {$properties = 'Model','Description','Released','Version','SupportedOS','Download'}
    else {$properties = 'Model','Released','Version','SupportedOS','Download'}

    # if the download location parameter has been used, allow passthru selection of desired downloads
    if ($DownloadLocation)
    {
        # display results in a grid view and set selected downloads variable to the user's selection
        $selectedDownloads = $sortedResults | Select-Object -Property $properties | Out-GridView -PassThru | Select-Object -ExpandProperty Download
        
        foreach ($selection in $selectedDownloads)
        {
            # split download selections that have more than one line into array
            $downloads = $selection -split [System.Environment]::NewLine
            
            foreach ($download in $downloads)
            {
                # get the filename substring
                $charIndex = ($download.LastIndexOf('/')) + 1
                $fileName = $download.Substring($charIndex,($download.Length) - $charIndex)
                
                # download the file
                Invoke-WebRequest -Uri $download -OutFile $DownloadLocation\$filename
            }    
        }
    }
    # else just display the results
    else
    {
        $sortedResults | Select-Object -Property $properties | Out-GridView
    }
}