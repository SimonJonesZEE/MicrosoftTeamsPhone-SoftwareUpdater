#region - Configuration (user editable)

<#
The manfacturers latest certified firmware (current) version must be configured here.
Teams certified devices: https://learn.microsoft.com/en-us/microsoftteams/devices/teams-ip-phones
#>

# Current Firmware
$currentfirmware = "122.15.0.157"

#endregion

###############################
#                             #
# DO NOT EDIT BELOW THIS LINE #
#                             #
###############################

#region - Preface
<#
 .SYNOPSIS
    Teams Device Updates
 .DESCRIPTION
    This script will update device software.
    This script was developed by Simon Jones (Jonesey).
    Version 1.5 - Updated on 08/08/2024 by Jonesey.
    Microsoft.Graph.Beta - v2.21.1
   
.SCRIPT REQUIREMENTS
    Module Microsoft.Graph.Beta
    
 .PARAMETER
    The manfacturers latest firmware version must be configured in the config.xml file under settings.
    Example for Yealink MP54/MP56/MP58 configuration: $currentfirmware = "122.15.0.157"
    Teams certified devices: https://learn.microsoft.com/en-us/microsoftteams/devices/teams-ip-phones

 .EXAMPLE
    .\updater_v1.5.ps1
#>
$host.ui.RawUI.WindowTitle = "GraphApp - Updater v1.5"
#endregion

#region - Startup
Cls
$path = split-path -parent $MyInvocation.MyCommand.Definition
Start-Transcript -Path "$path\logs\transcript.txt"   
#endregion

#region - Authentication
Function Auth {
    try {
        Connect-MgGraph -ContextScope Process -Scopes TeamworkDevice.ReadWrite.All
    }
        catch {
            ('Error - {0}' -f $_.Exception.Message) | Out-File "$path\logs\debugs.txt" -append
        }
}
#endregion

#region - Initialize
Auth
#endregion

#region - Text file selection
function Select-txtFile {
	    param([string]$Title="Please select the batch file",[string]$Directory=$("$path\data\"),[string]$Filter='Text Files (*.txt)|*.txt')
	    [System.Reflection.Assembly]::LoadWithPartialName("PresentationFramework") | Out-Null
	    $objForm = New-Object Microsoft.Win32.OpenFileDialog
	    $objForm.InitialDirectory = $Directory
	    $objForm.Filter = $Filter
	    $objForm.Title = $Title
	    $show = $objForm.ShowDialog()

    if ($show -eq $true) {
        [String]$txtFile = $objForm.FileName
        Write-Warning "The following file was selected: $($txtFile.Split("\")[$_.Length-2])"
		$continue = Read-Host "Press [C] to continue"

	if ($continue.ToLower() -eq "c") {
		Return $txtFile
		}
    }
    
    Return 0
}
#endregion

#region - Archive
function Archive {
    try {
        Write-Host 'Archiving files...'
        Stop-Transcript
        Get-ChildItem -Path "$path\data\*", "$path\exp\*", "$path\logs\*", "$path\tmp\*" | Compress-Archive -DestinationPath "$path\arc\$($sitecode)-Arc-$((Get-Date).ToString('dd-MM-yyyy')).zip" -Force -ErrorAction Stop
        Start-Transcript -Path "$path\logs\transcript.txt"
    }
        catch {
            ('Error: Archive Files - {0}' -f $_.Exception.Message) | Out-File "$path\logs\debugs.txt" -Append
        }
}
#endregion

#region - Batch File
function Batch-File {
    do {
        Cls
        $sitecode = Read-Host "Enter a Site Code (XXXXX)"
    }   while ($sitecode.Length -ne 5)

    if (-not([string]::IsNullOrEmpty($sitecode))) {
        $batchfile = Select-txtFile
    }
    if ($batchfile -ne '0') {
        $filename = ($($batchfile.Split("\")[$_.Length-2]))
    }
Options_Menu_GUI
}
#endregion

#region - Clear Debugs
function Clear-Debugs {
    if(Test-Path -Path "$path\logs\debugs.txt" -PathType Leaf) {
        Remove-Item "$path\logs\debugs.txt"
    }
}
#endregion

#region - Query Devices
function Query-Devices {
    if (Test-Path -Path "$path\exp\Query.csv" -PathType Leaf) {
        Clear-Content "$path\exp\Query.csv"
    } 
    if (Test-Path -Path "$path\exp\Offline.csv" -PathType Leaf) {
        Clear-Content "$path\exp\Offline.csv"
    } 
    if (-not([string]::IsNullOrEmpty($batchfile))) {
        $targetid = Get-Content "$batchfile"
        $userarrayq = $targetid | Measure
        $file = $userarrayq.count
        $i = 0
            foreach ($id in $targetid) {
                try {
                    $percentcomplete = ($i / $file) * 100; $i++
                    Write-Progress -Activity ('Querying device id: ' + $id) -Status ('file ' + $i + ' of ' + $file) -PercentComplete $percentcomplete
                    $healthstatus = Get-MgBetaTeamworkDevice -TeamworkDeviceId $id -ErrorAction Stop | ForEach {$_.HealthStatus}
                    $identity = Get-MgBetaTeamworkDevice -TeamworkDeviceId $id -ErrorAction Stop | ForEach {$_.Id} 
                    $displayname = Get-MgBetaTeamworkDevice -TeamworkDeviceId $id -ErrorAction Stop | ForEach {$_.CurrentUser.DisplayName} 
                    $firmwarecv = Get-MgBetaTeamworkDeviceHealth -TeamworkDeviceId $id -ErrorAction Stop | ForEach {$_.SoftwareUpdateHealth.FirmwareSoftwareUpdateStatus.CurrentVersion} 
                    $firmwareav = Get-MgBetaTeamworkDeviceHealth -TeamworkDeviceId $id -ErrorAction Stop | ForEach {$_.SoftwareUpdateHealth.FirmwareSoftwareUpdateStatus.AvailableVersion} 
                    $companyportalcv = Get-MgBetaTeamworkDeviceHealth -TeamworkDeviceId $id -ErrorAction Stop | ForEach {$_.SoftwareUpdateHealth.CompanyPortalSoftwareUpdateStatus.CurrentVersion} 
                    $companyportalav = Get-MgBetaTeamworkDeviceHealth -TeamworkDeviceId $id -ErrorAction Stop | ForEach {$_.SoftwareUpdateHealth.CompanyPortalSoftwareUpdateStatus.AvailableVersion} 
                    $teamscv = Get-MgBetaTeamworkDeviceHealth -TeamworkDeviceId $id -ErrorAction Stop | ForEach {$_.SoftwareUpdateHealth.TeamsClientSoftwareUpdateStatus.CurrentVersion} 
                    $teamsav = Get-MgBetaTeamworkDeviceHealth -TeamworkDeviceId $id -ErrorAction Stop | ForEach {$_.SoftwareUpdateHealth.TeamsClientSoftwareUpdateStatus.AvailableVersion} 

                    $params = [ordered] @{
                        "Healthstatus" = $healthstatus
                        "Identity" = $identity
                        "Displayname" = $displayname
                        "FirmwareCurrentVersion" = $firmwarecv
                        "FirmwareAvailableVersion" = $firmwareav
                        "CompanyPortalCurrentVersion" = $companyportalcv
                        "CompanyPortalAvailableVersion" = $companyportalav
                        "TeamsCurrentVersion" = $teamscv
                        "TeamsAvailableVersion" = $teamsav
                    }
                    $newrow = New-Object PSObject -Property $params 
                    $importtable += $newrow | Export-Csv "$path\exp\Query.csv" -Append -NoTypeInformation
                }
                    catch {
                        echo ('No Dice: {0}' -f $id) ('Error: Query Device ID - {0}' -f $_.Exception.Message) | Out-File "$path\logs\debugs.txt" -Append 
                    }
            }
    }
    if (Test-Path -Path "$path\tmp\numbers.txt" -PathType Leaf) {
        Clear-Content "$path\tmp\numbers.txt"
    }
    if (Test-Path -Path "$path\exp\Query.csv" -PathType Leaf) {
            $targetusers = Import-Csv "$path\exp\Query.csv"
            foreach ($user in $targetusers) {
                if ($user.Healthstatus -eq 'offline') {

                    $params = [ordered] @{
                        "Healthstatus" = $user.HealthStatus
                        "Identity" = $user.Identity
                        "Displayname" = $user.DisplayName
                    }
                    $newrow = New-Object PSObject -Property $params 
                    $importtable += $newrow | Export-Csv "$path\exp\Offline.csv" -Append -NoTypeInformation
                }
                if ($user.Healthstatus -eq 'nonUrgent') {
                    ($user.Identity | Select -Unique) | Out-File "$path\tmp\numbers.txt" -Append
                    $count = Get-Content "$path\tmp\numbers.txt"
                    $counter = ($count | Select -Unique).Count
                }else {
                    $count = Get-Content "$path\tmp\numbers.txt"
                    $counter = ($count | Select -Unique).Count
                }
            }
    }
Write-Progress -Activity "Sleep" -Completed
Options_Menu_GUI
}
#endregion

#region - Device Updates
function Device-Updates {
    if ($counter -gt 0) {
        $confirm = Read-Host "Proceed with device updates? [Y] Yes [N] No"
        Cls
    }
    if ($confirm -eq 'y' -and $up -ne 'update') {
        $targetid = Get-Content "$path\tmp\numbers.txt"
        $userarrayq = $targetid | Measure
        $file = $userarrayq.count
        $i = 0
            try {
                foreach ($user in $targetid) {
                $health = (Get-MgBetaTeamworkDeviceHealth -TeamworkDeviceId $user)
                $uid = (Get-MgBetaTeamworkDevice -TeamworkDeviceId $user)
                $percentcomplete = ($i / $file) * 100; $i++
                Write-Progress -Activity ('Queueing device updates on id: ' + $user) -Status ('file ' + $i + ' of ' + $file) -PercentComplete $percentcomplete
                #this if will check to see if the device's firmware is above the ms current version and downgrade the device's firmware.
                if ($health.softwareupdatehealth.FirmwareSoftwareUpdateStatus.AvailableVersion -eq $null -and $health.softwareupdatehealth.FirmwareSoftwareUpdateStatus.CurrentVersion -ne $currentfirmware) {
                    Update-MgBetaTeamworkDeviceSoftware -TeamworkDeviceId $user -SoftwareType "firmware" -SoftwareVersion $currentfirmware -ErrorAction Stop
                    $sid = (Get-MgBetaTeamworkDeviceOperation -TeamworkDeviceId $user -ErrorAction Stop | ForEach {$_.Id} | Select -Index '0')
  
                    $params = [ordered] @{
                        "Healthstatus" = $uid.HealthStatus
                        "Identity" = $uid.Id
                        "Displayname" = $uid.CurrentUser.DisplayName
                        "SID" = $sid
                        "Softwaretype" = 'Firmware'
                    }
                    $newrow = New-Object PSObject -Property $params 
                    $importtable += $newrow | Export-Csv "$path\tmp\updates.csv" -Append -NoTypeInformation
                    $up = 'update'
                    #this elseif will check to see if the device's firmware requires an upgrade.
                    }elseif ($health.softwareupdatehealth.FirmwareSoftwareUpdateStatus.AvailableVersion -ne $null `
                             -and $health.softwareupdatehealth.FirmwareSoftwareUpdateStatus.AvailableVersion -ne $health.softwareupdatehealth.FirmwareSoftwareUpdateStatus.CurrentVersion) {
                        Update-MgBetaTeamworkDeviceSoftware -TeamworkDeviceId $user -SoftwareType "firmware" -SoftwareVersion $health.softwareupdatehealth.FirmwareSoftwareUpdateStatus.AvailableVersion -ErrorAction Stop
                        $sid = (Get-MgBetaTeamworkDeviceOperation -TeamworkDeviceId $user -ErrorAction Stop | ForEach {$_.Id} | Select -Index '0')

                        $params = [ordered] @{
                            "Healthstatus" = $uid.HealthStatus
                            "Identity" = $uid.Id
                            "Displayname" = $uid.CurrentUser.DisplayName
                            "SID" = $sid
                            "Softwaretype" = 'Firmware'
                        }
                        $newrow = New-Object PSObject -Property $params 
                        $importtable += $newrow | Export-Csv "$path\tmp\updates.csv" -Append -NoTypeInformation
                        $up = 'update'
                    }

                if ($health.softwareupdatehealth.CompanyPortalSoftwareUpdateStatus.AvailableVersion -eq $null) {
                    #this elseif will check to see if the device's company portal software requires an upgrade.
                    }elseif ($health.softwareupdatehealth.CompanyPortalSoftwareUpdateStatus.AvailableVersion -ne $health.softwareupdatehealth.CompanyPortalSoftwareUpdateStatus.CurrentVersion) {
                        Update-MgBetaTeamworkDeviceSoftware -TeamworkDeviceId $user -SoftwareType "companyPortal" -SoftwareVersion $health.softwareupdatehealth.CompanyPortalSoftwareUpdateStatus.AvailableVersion -ErrorAction Stop
                        $sid = (Get-MgBetaTeamworkDeviceOperation -TeamworkDeviceId $user -ErrorAction Stop | ForEach {$_.Id} | Select -Index '0')

                        $params = [ordered] @{
                            "Healthstatus" = $uid.HealthStatus
                            "Identity" = $uid.Id
                            "Displayname" = $uid.CurrentUser.DisplayName
                            "SID" = $sid
                            "Softwaretype" = 'Company Portal'
                        }
                        $newrow = New-Object PSObject -Property $params 
                        $importtable += $newrow | Export-Csv "$path\tmp\updates.csv" -Append -NoTypeInformation
                        $up = 'update'
                    }

                if ($health.softwareupdatehealth.TeamsClientSoftwareUpdateStatus.AvailableVersion -eq $null) {
                    #this elseif will check to see if the device's teams app software requires an upgrade.
                    }elseif ($health.softwareupdatehealth.TeamsClientSoftwareUpdateStatus.AvailableVersion -ne $health.softwareupdatehealth.TeamsClientSoftwareUpdateStatus.CurrentVersion) {
                        Update-MgBetaTeamworkDeviceSoftware -TeamworkDeviceId $user -SoftwareType "teamsClient" -SoftwareVersion $health.softwareupdatehealth.TeamsClientSoftwareUpdateStatus.AvailableVersion -ErrorAction Stop
                        $sid = (Get-MgBetaTeamworkDeviceOperation -TeamworkDeviceId $user -ErrorAction Stop | ForEach {$_.Id} | Select -Index '0')
                     
                        $params = [ordered] @{
                            "Healthstatus" = $uid.HealthStatus
                            "Identity" = $uid.Id
                            "Displayname" = $uid.CurrentUser.DisplayName
                            "SID" = $sid
                            "Softwaretype" = 'Teams'
                        }
                        $newrow = New-Object PSObject -Property $params 
                        $importtable += $newrow | Export-Csv "$path\tmp\updates.csv" -Append -NoTypeInformation
                        $up = 'update'
                    }
                }
             }
                catch {
                    echo ('No Dice: {0}' -f $user.DisplayName) ('Error: Software Updates - {0}' -f $_.Exception.Message) | Out-File "$path\logs\debugs.txt" -Append
                }
    }
Write-Progress -Activity "Sleep" -Completed
Options_Menu_GUI
}
#endregion

#region - Update Status
function Update-Status {
    if (Test-Path -Path "$path\exp\Status.csv" -PathType Leaf) {
        Clear-Content "$path\exp\Status.csv"
    }
    if (Test-Path -Path "$path\tmp\updates.csv" -PathType Leaf) {
        $targetusers = Import-Csv "$path\tmp\updates.csv"
        $userarrayq = $targetusers | Measure
        $file = $userarrayq.count
        $i = 0
            foreach ($user in $targetusers) {
                Try {
                    $percentcomplete = ($i / $file) * 100; $i++
                    Write-Progress -Activity ('Checking device updates on... ' + $user.Displayname) -Status ('file ' + $i + ' of ' + $file) -PercentComplete $percentcomplete
                    $status = (Get-MgBetaTeamworkDeviceOperation -TeamworkDeviceId $user.Identity -ErrorAction Stop | Where-Object {$_.Id -eq $user.SID} | ForEach {$_.Status}) 
                    $healthstatus = (Get-MgBetaTeamworkDevice -TeamworkDeviceId $user.Identity -ErrorAction Stop | ForEach {$_.HealthStatus})

                    $params = [ordered] @{
                        "Software type" = $user.Softwaretype
                        "Status" = $status
                        "Identity" = $user.Identity
                        "Healthstatus" = $healthstatus
                        "Display name" = $user.Displayname  
                    }
                    $newrow = New-Object PSObject -Property $params 
                    $importtable += $newrow | Export-Csv "$path\exp\Status.csv" -Append -NoTypeInformation
                }
                    catch {
                        echo ('No Dice: {0}' -f $user.UPN) ('Error: Device Update Status - {0}' -f $_.Exception.Message) | Out-File "$path\logs\debugs.txt" -Append
                    }
            }
    }
    if (Test-Path -Path "$path\tmp\numbers.txt" -PathType Leaf) {
        Clear-Content "$path\tmp\numbers.txt"
    }
    if (Test-Path -Path "$path\exp\Status.csv" -PathType Leaf) {
            $targetusers = Import-Csv "$path\exp\Status.csv"
            foreach ($user in $targetusers) {
                if ($user.Healthstatus -eq 'nonUrgent') {
                    ($user.Identity | Select -Unique) | Out-File "$path\tmp\numbers.txt" -Append
                    $count = Get-Content "$path\tmp\numbers.txt"
                    $counter = ($count | Select -Unique).Count
                }else {
                    $count = Get-Content "$path\tmp\numbers.txt"
                    $counter = ($count | Select -Unique).Count
                }
            }
    }
Write-Progress -Activity "Sleep" -Completed
Options_Menu_GUI
}
#endregion

#region - Restart Devices
function Restart-Devices {
$restarts = Select-txtFile
Cls
$confirm = Read-Host "Proceed with device restart? [Y] Yes [N] No"
Cls
    if ($confirm -eq 'Y') {
        $targetid = Get-Content "$restarts"
        $userarrayq = $targetid | Measure
        $file = $userarrayq.count
        $i = 0
        try {
            foreach ($id in $targetid) {
                $percentcomplete = ($i / $file) * 100; $i++
                Write-Progress -Activity ('Restarting device... ' + $id) -Status ('file ' + $i + ' of ' + $file) -PercentComplete $percentcomplete
                Restart-MgBetaTeamworkDevice -TeamworkDeviceId $id -ErrorAction Stop
            }
        }   
            catch {
                echo ('No Dice: {0}' -f $user) ('Error: Restart Devices - {0}' -f $_.Exception.Message) | Out-File "$path\logs\debugs.txt" -Append
            }
    }
Write-Progress -Activity "Sleep" -Completed
}
#endregion

#region - Quit
function Quit {
$confirm = Read-Host "This will remove session files. Continue? [Y] Yes [N] No"
    if ($confirm -eq 'Y') {
        Stop-Transcript
        Disconnect-MgGraph
        Remove-Item "$path\exp\*", "$path\logs\debugs.txt", "$path\tmp\*"
    }
    if ($confirm -eq 'N') {
        Stop-Transcript
        Disconnect-MgGraph 
    }
}
#endregion

#region - Sub Functions
function Option-A {
Cls
Archive
}

function Option-B { 
Cls
Batch-File
}

function Option-C {
Cls
Clear-Debugs
}

function Option-1 {
Cls
Query-Devices
}

function Option-2 {
Cls
Device-Updates
}

function Option-3 {
Cls
Update-Status
}

function Option-R {
Cls
Restart-Devices
}

function Option-Q {
Cls
Quit
}
#endregion

#region - Menu GUI
function Options_Menu_GUI {
do
{
Cls
if (([string]::IsNullOrEmpty($sitecode))){
    $sitecode = 'XXXXX'
}     
    Write-Host -Object '***************************'
    Write-Host -Object '*                         *'
    Write-Host -Object "* $sitecode - Nuffield Health *"
    Write-Host -Object '*                         *'
    Write-Host -Object '***************************'
    Write-Host -Object ''
    Write-Host -Object 'Choose an option' -ForegroundColor Yellow
    Write-Host -Object ''
    Write-Host -Object 'A.   Archive (current)' 
    Write-Host -Object ''
    Write-Host -Object "B.   Batch File ($filename)"
    Write-Host -Object ''
    Write-Host -Object 'C.   Clear Debugs'
    Write-Host -Object ''
    Write-Host -Object '1.   Query Devices'
    Write-Host -Object ''
    Write-Host -Object "2.   Device Updates ($counter)"
    Write-Host -Object ''
    Write-Host -Object '3.   Update Status'
    Write-Host -Object ''
    Write-Host -Object 'R.   Restart Devices'
    Write-Host -Object ''
    Write-Host -Object 'Q.   Quit'
    Write-Host -Object ''
    $Menu = Read-Host -Prompt '(A, B, C, R, 1-3 or Q to Quit)'

    Switch ($Menu) {
            A
        {
            Option-A
        }
            B
        {
            Option-B
        }
            C
        {
            Option-C
        }
            1
        {
            Option-1
        }
            2
        {
            Option-2
        }
            3
        {
            Option-3
        }
            R
        {
            Option-R
        }
            Q
        {
            Option-Q
            Cls
            Exit
        }
    }
}
    until ($Menu -eq 'q')
}
Options_Menu_GUI
#endregion