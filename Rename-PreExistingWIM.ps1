<#
    Author: Kevin Johnston
    Date: 10/5/2017
    
    Checks for a preexisting WIM file in the captures folder with the
    same name as the one to be captured and renames it. This prevents
    the current capture from being added to the existing WIM. 
#>

$TSEnv = New-Object -ComObject Microsoft.SMS.TSEnvironment
$WIMFileName = $TSEnv.Value("BackupFile")
$WIMPath = $TSEnv.Value("ComputerBackupLocation") + "\" + $WIMFileName

if (Test-Path -Path $WIMPath)
{
    $WIMNewName = $WIMFileName + ".bak" + (Get-Date -Format yyyyMMdd-HHmmss)
    Rename-Item -Path $WIMPath -NewName $WIMNewName
}