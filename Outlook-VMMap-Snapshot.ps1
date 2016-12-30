# Author: Kevin Johnston
# Date:   April 7, 2016
#
# This script performs the following actions:
#
# 1. Opens/Displays/Renders and closes an Outlook message for a defined number of cycles
# 2. Runs VMMap at a defined cycle interval to generate .mmp (virtual memory snapshot) files
# 3. Parses the .mmp XML content to find the count of 4KB private data allocations as well as unusable and non-free virtual memory
# 4. Outputs cycle progress and VMMap information to the console
#
# Tested with Outlook 2010*, 2013, and 2016
# *Please see the note on line 34 regarding method change for Outlook 2010   


$cycles = 500                           # The maximum number of open/close message cycles
$vmmapinterval = 50                     # The cycle interval at which VMMap will run and generate a .mmp file
$vmmapfolder = "C:\Temp\vmmap"          # The location of VMMap.exe and the save location for .mmp files
$mailboxname = "email@yourcompany.com"  # The desired Outlook mailbox Name (Likely your email address)
$mailfoldername = "Inbox"               # The desired mailbox folder name 

# Create the Outlook COM object and get the messaging API namespace
$outlook = New-Object -ComObject Outlook.Application 
$namespace = $outlook.GetNamespace("MAPI")

# Create the mailbox and mailfolder objects
$mailbox = $namespace.Folders | Where-Object {$_.Name -eq $mailboxname}
$mailfolder = $mailbox.Folders.Item($mailfoldername)

# Display the Outlook main window
$explorer = $mailfolder.GetExplorer()
$explorer.Display()

# Create the message object
$message = $mailfolder.Items.GetLast() # Change to .GetFirst() method if using Outlook 2010, otherwise .Close() method will not work

# Add the assembly needed to create the OlInspectorClose object for the .Close() method
Add-Type -Assembly "Microsoft.Office.Interop.Outlook"
$discard = [Microsoft.Office.Interop.Outlook.OlInspectorClose]::olDiscard

#-------------------------------------------------------------------------------------------------------------------------------------
# Execute the above code first, wait for the Outlook window to display, and reposition it if necessary before executing the below code
#-------------------------------------------------------------------------------------------------------------------------------------

for ($i = 1; $i -lt ($cycles + 1) ; $i++)
{ 
    # Open the message then close and discard changes
    $message.Display()
    $message.Close($discard)

    Write-Progress -Activity "Working..." -Status "$i of $cycles cycles complete" -PercentComplete (($i / $cycles) * 100)

    if ($i % $vmmapinterval -eq 0)
    {
        # Run VMMap map with the necessary command line options and generate .mmp file
        Start-Process -Wait -FilePath $vmmapfolder\vmmap.exe -ArgumentList "-accepteula -p outlook.exe outputfile $vmmapfolder\outlook$i.mmp" -WindowStyle Hidden

        # Get .mmp file content as XML
        [xml]$vmmap = Get-Content $vmmapfolder\outlook$i.mmp
        $regions = $vmmap.root.Snapshots.Snapshot.MemoryRegions.Region
        
        # Get Count of 4KB private data allocations
        $privdata4k = ($regions | Where-Object {($_.Type -eq "Private Data") -and ($_.Size -eq "4096")}).Count
        
        # Get Unusable and non-free virtual memory totals 
        $unusablevm = ((($regions | Where-Object {$_.Type -eq "Unusable"}).Size | Measure-Object -Sum).Sum / 1MB)
        $nonfreevm = ((($regions | Where-Object {$_.Type -ne "Free"}).Size | Measure-Object -Sum).Sum / 1GB)
        
        # Round results to two decimal places
        $unusablevmrounded = [math]::Round($unusablevm,2)
        $nonfreevmrounded = [math]::Round($nonfreevm,2)

        Write-Output "-----------------------------------------------------------------------"
        Write-Output "   $privdata4k 4KB Private Data Allocations and"
        Write-Output "   $unusablevmrounded MBs of Unusable Memory After $i Open/Close Cycles"
        Write-Output "   $nonfreevmrounded GB of 2GB Virtual Memory Limit Reached"
        Write-Output "-----------------------------------------------------------------------"
        
    }
}
