<#PSScriptInfo

.VERSION 2.0

.GUID
cf43c4ef-a648-4f67-854f-017b7394ec3c

.AUTHOR
Simon Jones

.COMPANYNAME
SACOMS

.COPYRIGHT
(c) 2024 Simon Jones. All rights reserved.

.TAGS

.LICENSEURI

.PROJECTURI
https://github.com/SimonJonesZEE/MicrosoftTeamsPhone-SoftwareUpdater

.ICONURI

.EXTERNALMODULEDEPENDENCIES

.REQUIREDSCRIPTS

.EXTERNALSCRIPTDEPENDENCIES

.RELEASENOTES

.PRIVATEDATA

#>

<#

.DESCRIPTION
Microsoft Teams Phone Software Updater

#>
Param()

#################################
#                               #
#  DO NOT EDIT BELOW THIS LINE  #
#                               #
#################################

#region - STARTUP

# This will set the session windows title bar with the name of the script
$host.ui.RawUI.WindowTitle = "Teams Phone Software Updater v1.0"
# screen wash
Clear-Host
# set the path variable
$path = "$PSScriptRoot"
# create folder structure
if ((Test-Path -Path "$path\arc") -eq $false) {
    New-Item -Path "$path\" -Name "arc" -ItemType "directory"
}
if ((Test-Path -Path "$path\data") -eq $false) {
    New-Item -Path "$path\" -Name "data" -ItemType "directory"
}
if ((Test-Path -Path "$path\logs") -eq $false) {
    New-Item -Path "$path\" -Name "logs" -ItemType "directory"
}
if ((Test-Path -Path "$path\tmp") -eq $false) {
    New-Item -Path "$path\" -Name "tmp" -ItemType "directory"
}
# start the transcript
Start-Transcript -Path "$path\logs\transcript.txt"

# Connect the Microsoft.Graph PowerShell module
Function Auth {
    try {
        Connect-MgGraph -ContextScope Process -Scopes TeamworkDevice.ReadWrite.All -ErrorAction Stop
    }
        catch {
        Exit
        ('Error - {0}' -f $_.Exception.Message) | Out-File "$path\logs\debug.txt" -append
        }
}

# This will call the Auth function when this script is run
Auth

#endregion

#region - FUNCTIONS

# This will prompt the user to select a text file which contains one or more TeamworkDeviceId's
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
        $check = Get-Content $txtFile
	if (($continue.ToLower() -eq "c") -and (-not[String]::IsNullOrWhiteSpace($check))) {
        Return $txtFile
	    }else {
            Write-Error 'No data in selected text file.'
        }
    }
    Return 0
}

# Archives the session files and stores them in the '$path\Arc\' folder
function Archive {
    try {
        Stop-Transcript
        Get-ChildItem -Path "$path\tmp\*", "$path\logs\*" | Compress-Archive -DestinationPath "$path\arc\$($customer)-$($site)-$((Get-Date).ToString('dd-MM-yyyy')).zip" -Force -ErrorAction Stop
        Start-Transcript -Path "$path\logs\transcript.txt"
    }
        catch {
            ('Error: Archive Reports - {0}' -f $_.Exception.Message) | Out-File "$path\logs\debug.txt" -Append
        }
}

# Prompt the user to input a site name or code for reference
function Batch_File {
$confirm = Read-Host "Are you sure you want to load a new file? [Y] Yes [N] No"
    if ($confirm -eq 'Y') {
        do {
            Clear-Host
            $site = Read-Host "Enter a site name or code"
            $displaysite = '(' + $site + ')'
        }
        while ($site -like '')
        $batchfile = Select-txtFile
        if ($batchfile -ne '0' -and $null -ne $batchfile) {
            $filename = ($($batchfile.Split("\")[$_.Length-2]))
            $displayfilename = '(' + $filename + ')'
            # if a new file is selected clear variable and tmp
            $displaycounter = $null
            $displaycounter2 = $null
            $ids = $null
            $flag = $null
            $flag2 = $null
            Remove-Item "$path\tmp\*"
        }else {
            # if no file is selected clear variable and tmp
            $batchfile = $null
            $displaysite = $null
            $displayfilename = $null
            $displaycounter = $null
            $displaycounter2 = $null
            $ids = $null
            $flag = $null
            $flag2 = $null
            Remove-Item "$path\tmp\*"
        }
    Options_Menu_GUI
    }
}

# Clear the debug.txt file
function Clear_debug {
    if(Test-Path -Path "$path\logs\debug.txt" -PathType Leaf) {
        Remove-Item "$path\logs\debug.txt"
    }
}

# Verify TeamworkDeviceId's against the batch file and export data to a Csv file so it can be imported for computation and user analysis
function Analyze_file {
    if ($null -ne $batchfile -and (Test-Path -LiteralPath "$path\tmp\did.csv") -eq $false) {
        $targetid = Get-Content "$batchfile"
        $userarrayq = $targetid | Measure-Object
        $file = $userarrayq.count
        $i = 0
        try {
            foreach ($id in $targetid) {
                $percentcomplete = ($i / $file) * 100; $i++
                Write-Progress -Activity ('Analyzing TeamworkDeviceId: ' + $id) -Status ('file ' + $i + ' of ' + $file) -PercentComplete $percentcomplete
                $hs = Get-MgBetaTeamworkDevice -TeamworkDeviceId $id -ErrorAction Stop | ForEach-Object {$_.HealthStatus}
                $id = Get-MgBetaTeamworkDevice -TeamworkDeviceId $id -ErrorAction Stop | ForEach-Object {$_.Id}
                $dn = Get-MgBetaTeamworkDevice -TeamworkDeviceId $id -ErrorAction Stop | ForEach-Object {$_.CurrentUser.DisplayName}
                $sw = Get-MgBetaTeamworkDeviceHealth -TeamworkDeviceId $id -ErrorAction Stop
                $fw = $sw.softwareupdatehealth.FirmwareSoftwareUpdateStatus
                $cp = $sw.softwareupdatehealth.CompanyPortalSoftwareUpdateStatus
                $tm = $sw.softwareupdatehealth.TeamsClientSoftwareUpdateStatus
                # create and add data to did
                $params = [ordered] @{
                    "HealthStatus" = $hs | Where-Object {$fw.SoftwareFreshness -ne 'latest'}
                    "Id" = $id
                    "DisplayName" = $dn
                    "SoftwareType" = 'FirmwareUpdate'
                    "Status" = ''
                    "SID" = ''
                    "AvailableVersion" = $fw.AvailableVersion
                    "CurrentVersion" = $fw.CurrentVersion
                    "SoftwareFreshness" = $fw.SoftwareFreshness
                    "Sync" = ''
                }
                $newrow = New-Object PSObject -Property $params
                $importtable += $newrow | Export-Csv "$path\tmp\did.csv" -Append -NoTypeInformation
                # add data to did
                $params = [ordered] @{
                    "HealthStatus" = $hs | Where-Object {$cp.SoftwareFreshness -ne 'latest'}
                    "Id" = $id
                    "DisplayName" = $dn
                    "SoftwareType" = 'CompanyPortalUpdate'
                    "Status" = ''
                    "SID" = ''
                    "AvailableVersion" = $cp.AvailableVersion
                    "CurrentVersion" = $cp.CurrentVersion
                    "SoftwareFreshness" = $cp.SoftwareFreshness
                    "Sync" = ''
                }
                $newrow = New-Object PSObject -Property $params
                $importtable += $newrow | Export-Csv "$path\tmp\did.csv" -Append -NoTypeInformation
                # add data to did
                $params = [ordered] @{
                    "HealthStatus" = $hs | Where-Object {$tm.SoftwareFreshness -ne 'latest'}
                    "Id" = $id
                    "DisplayName" = $dn
                    "SoftwareType" = 'TeamsClientUpdate'
                    "Status" = ''
                    "SID" = ''
                    "AvailableVersion" = $tm.AvailableVersion
                    "CurrentVersion" = $tm.CurrentVersion
                    "SoftwareFreshness" = $tm.SoftwareFreshness
                    "Sync" = ''
                 }
                 $newrow = New-Object PSObject -Property $params
                 $importtable += $newrow | Export-Csv "$path\tmp\did.csv" -Append -NoTypeInformation
            }
        }
        catch {Write-Output ('No Dice: {0}' -f $id) ('Error: Analyze File - {0}' -f $_.Exception.Message) | Out-File "$path\logs\debug.txt" -Append}
    $num = @()
    $dev = @()
    $targetusers = Import-Csv "$path\tmp\did.csv"
        foreach ($user in $targetusers) {
            $num += ($user.Id | Select-Object -Unique | Where-Object {$user.HealthStatus -eq 'nonUrgent'})
            $dev += ($user.Id | Select-Object -Unique)
        }
    $counter = ($num | Select-Object -Unique).Count
    $displaycounter = '(' + $counter + ')'
    $devices = ($dev | Select-Object -Unique).Count
    $ids = '- ' + $devices + ' x device(s)'
    $flag = $false
    }

Write-Progress -Activity "Sleep" -Completed
Options_Menu_GUI
}

# Update software
function Update_software {
    if ($null -ne $batchfile -and $counter -gt 0 -and (Test-Path -LiteralPath "$path\tmp\did.csv") -eq $true) {
        $targetid = Import-Csv "$path\tmp\did.csv"
        $csv = "$path\tmp\did.csv"
        $ofile = New-Object System.IO.FileInfo $csv
    try {
        # check if did.csv file is closed before continuing
        $ostream = $ofile.Open([System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
        if ($ostream) {$ostream.Close()}
        try {
            foreach ($user in $targetid | Where-Object {$_.HealthStatus -eq 'nonUrgent'}) {
                $uid = Get-MgBetaTeamworkDevice -TeamworkDeviceId $user.Id -ErrorAction Stop
                $sw = Get-MgBetaTeamworkDeviceHealth -TeamworkDeviceId $user.Id -ErrorAction Stop
                $fw = $sw.softwareupdatehealth.FirmwareSoftwareUpdateStatus
                $cp = $sw.softwareupdatehealth.CompanyPortalSoftwareUpdateStatus
                $tm = $sw.softwareupdatehealth.TeamsClientSoftwareUpdateStatus
                $ops = Get-MgBetaTeamworkDeviceOperation -TeamworkDeviceId $user.Id -ErrorAction Stop | Where-Object {$_.Id -eq $user.SID}
                # handle with a SID if conditions are met
                if ($ops.Status -eq 'Queued' -or $ops.Status -eq 'InProgress') {
                    Write-Progress ('Software update ' + $ops.Status + ' on: ' + $user.DisplayName)
                    # add data to did
                    $user.HealthStatus = $uid.HealthStatus
                    $user.Status = $ops.Status
                    $targetid | Export-Csv "$path\tmp\did.csv" -NoTypeInformation
                }elseif ($ops.Status -eq 'Failed') {
                    # re-attempt firmware update
                    if ($user.SoftwareType -eq 'FirmwareUpdate') {
                        Write-Progress -Activity ('Retrying Software update on: ' + $user.DisplayName)
                        Update-MgBetaTeamworkDeviceSoftware -TeamworkDeviceId $user.Id -SoftwareType "firmware" -SoftwareVersion $user.AvailableVersion -ErrorAction Stop
                        $sid = (Get-MgBetaTeamworkDeviceOperation -TeamworkDeviceId $user.Id -ErrorAction Stop | ForEach-Object {$_.Id} | Select-Object -Index '0')
                        # add data to did
                        $user.SID = $sid
                        $user.Status = 'Queued'
                        $targetid | Export-Csv "$path\tmp\did.csv" -NoTypeInformation
                    }
                    # re-attempt company portal update
                    if ($user.SoftwareType -eq 'CompanyPortalUpdate') {
                        Write-Progress -Activity ('Retrying Software update on: ' + $user.DisplayName)
                        Update-MgBetaTeamworkDeviceSoftware -TeamworkDeviceId $user.Id -SoftwareType "companyPortal" -SoftwareVersion $user.AvailableVersion -ErrorAction Stop
                        $sid = (Get-MgBetaTeamworkDeviceOperation -TeamworkDeviceId $user.Id -ErrorAction Stop | ForEach-Object {$_.Id} | Select-Object -Index '0')
                        # add data to did
                        $user.SID = $sid
                        $user.Status = 'Queued'
                        $targetid | Export-Csv "$path\tmp\did.csv" -NoTypeInformation
                    }
                    # re-attempt teams update
                    if ($user.SoftwareType -eq 'TeamsClientUpdate') {
                        Write-Progress -Activity ('Retrying Software update on: ' + $user.DisplayName)
                        Update-MgBetaTeamworkDeviceSoftware -TeamworkDeviceId $user.Id -SoftwareType "teamsClient" -SoftwareVersion $user.AvailableVersion -ErrorAction Stop
                        $sid = (Get-MgBetaTeamworkDeviceOperation -TeamworkDeviceId $user.Id -ErrorAction Stop | ForEach-Object {$_.Id} | Select-Object -Index '0')
                        # add data to did
                        $user.SID = $sid
                        $user.Status = 'Queued'
                        $targetid | Export-Csv "$path\tmp\did.csv" -NoTypeInformation
                    }
                }elseif ($ops.Status -eq 'Successful') {
                    Write-Progress -Activity ('Please wait...')
                    # check for step firmware from previous update
                   if ($user.SoftwareType -eq 'FirmwareUpdate' -and $user.Sync -eq $fw.CurrentVersion) {
                        # prep for step
                        if ($fw.SoftwareFreshness -eq 'updateAvailable') {
                            $user.HealthStatus = $uid.HealthStatus
                            $user.Status = ''
                            $user.SID = ''
                            $user.AvailableVersion = $fw.AvailableVersion
                            $user.CurrentVersion = $fw.CurrentVersion
                            $user.SoftwareFreshness = $fw.SoftwareFreshness
                            $user.Sync = ''
                            $targetid | Export-Csv "$path\tmp\did.csv" -NoTypeInformation
                        # clean up previous update
                        }else {
                            $user.HealthStatus = ''
                            $user.Status = $ops.Status
                            $user.SID = ''
                            $user.AvailableVersion = $fw.AvailableVersion
                            $user.CurrentVersion = $fw.CurrentVersion
                            $user.SoftwareFreshness = $fw.SoftwareFreshness
                            $user.Sync = ''
                            $targetid | Export-Csv "$path\tmp\did.csv" -NoTypeInformation
                        }
                    # wait for sync
                    }elseif ($user.SoftwareType -eq 'FirmwareUpdate' -and $user.Sync -ne $fw.CurrentVersion) {
                    $user.Status = $ops.Status
                    $targetid | Export-Csv "$path\tmp\did.csv" -NoTypeInformation
                    }
                    # check for step company portal from previous update
                     if ($user.SoftwareType -eq 'CompanyPortalUpdate' -and $user.Sync -eq $cp.CurrentVersion) {
                        # prep for step
                        if ($cp.SoftwareFreshness -eq 'updateAvailable') {
                            $user.HealthStatus = $uid.HealthStatus
                            $user.Status = ''
                            $user.SID = ''
                            $user.AvailableVersion = $cp.AvailableVersion
                            $user.CurrentVersion = $cp.CurrentVersion
                            $user.SoftwareFreshness = $cp.SoftwareFreshness
                            $user.Sync = ''
                            $targetid | Export-Csv "$path\tmp\did.csv" -NoTypeInformation
                        # clean up previous update
                        }else {
                            $user.HealthStatus = ''
                            $user.Status = $ops.Status
                            $user.SID = ''
                            $user.AvailableVersion = $cp.AvailableVersion
                            $user.CurrentVersion = $cp.CurrentVersion
                            $user.SoftwareFreshness = $cp.SoftwareFreshness
                            $user.Sync = ''
                            $targetid | Export-Csv "$path\tmp\did.csv" -NoTypeInformation
                        }
                    # wait for sync
                    }elseif ($user.SoftwareType -eq 'CompanyPortalUpdate' -and $user.Sync -ne $cp.CurrentVersion) {
                    $user.Status = $ops.Status
                    $targetid | Export-Csv "$path\tmp\did.csv" -NoTypeInformation
                    }
                    # check for step teams client from previous update
                    if ($user.SoftwareType -eq 'TeamsClientUpdate' -and $user.Sync -eq $tm.CurrentVersion) {
                        if ($tm.SoftwareFreshness -eq 'updateAvailable') {
                            $user.HealthStatus = $uid.HealthStatus
                            $user.Status = ''
                            $user.SID = ''
                            $user.AvailableVersion = $tm.AvailableVersion
                            $user.CurrentVersion = $tm.CurrentVersion
                            $user.SoftwareFreshness = $tm.SoftwareFreshness
                            $user.Sync = ''
                            $targetid | Export-Csv "$path\tmp\did.csv" -NoTypeInformation
                        # clean up previous update
                        }else {
                            $user.HealthStatus = ''
                            $user.Status = $ops.Status
                            $user.SID = ''
                            $user.AvailableVersion = $tm.AvailableVersion
                            $user.CurrentVersion = $tm.CurrentVersion
                            $user.SoftwareFreshness = $tm.SoftwareFreshness
                            $user.Sync = ''
                            $targetid | Export-Csv "$path\tmp\did.csv" -NoTypeInformation
                        }
                    # wait for sync
                    }elseif ($user.SoftwareType -eq 'TeamsClientUpdate' -and $user.Sync -ne $tm.CurrentVersion) {
                    $user.Status = $ops.Status
                    $targetid | Export-Csv "$path\tmp\did.csv" -NoTypeInformation
                    }
                        }else {
                            # check if the firmware requires an update
                            Write-Progress -Activity ('Update software on: ' + $user.DisplayName)
                            if ($null -ne $fw.AvailableVersion -and $fw.AvailableVersion -ne $fw.CurrentVersion | Where-Object {$user.SoftwareType -eq 'FirmwareUpdate'}) {
                                Update-MgBetaTeamworkDeviceSoftware -TeamworkDeviceId $user.Id -SoftwareType "firmware" -SoftwareVersion $fw.AvailableVersion -ErrorAction Stop
                                $sid = (Get-MgBetaTeamworkDeviceOperation -TeamworkDeviceId $user.Id -ErrorAction Stop | ForEach-Object {$_.Id} | Select-Object -Index '0')
                                # add data to did
                                $user.SID = $sid
                                $user.SoftwareType = 'FirmwareUpdate'
                                $user.Status = 'Queued'
                                $user.AvailableVersion = $fw.AvailableVersion
                                $user.CurrentVersion = $fw.CurrentVersion
                                $user.SoftwareFreshness = $fw.SoftwareFreshness
                                $user.Sync = $fw.AvailableVersion
                                $targetid | Export-Csv "$path\tmp\did.csv" -NoTypeInformation
                            }
                            # check if the company portal requires an update
                            if ($null -ne $cp.AvailableVersion -and $cp.AvailableVersion -ne $cp.CurrentVersion | Where-Object {$user.SoftwareType -eq 'CompanyPortalUpdate'}) {
                                Update-MgBetaTeamworkDeviceSoftware -TeamworkDeviceId $user.Id -SoftwareType "companyPortal" -SoftwareVersion $cp.AvailableVersion -ErrorAction Stop
                                $sid = (Get-MgBetaTeamworkDeviceOperation -TeamworkDeviceId $user.Id -ErrorAction Stop | ForEach-Object {$_.Id} | Select-Object -Index '0')
                                # add data to did
                                $user.SID = $sid
                                $user.SoftwareType = 'CompanyPortalUpdate'
                                $user.Status = 'Queued'
                                $user.AvailableVersion = $cp.AvailableVersion
                                $user.CurrentVersion = $cp.CurrentVersion
                                $user.SoftwareFreshness = $cp.SoftwareFreshness
                                $user.Sync = $cp.AvailableVersion
                                $targetid | Export-Csv "$path\tmp\did.csv" -NoTypeInformation
                            }
                            # check if the teams app requires an update
                            if ($null -ne $tm.AvailableVersion -and $tm.AvailableVersion -ne $tm.CurrentVersion | Where-Object {$user.SoftwareType -eq 'TeamsClientUpdate'}) {
                                Update-MgBetaTeamworkDeviceSoftware -TeamworkDeviceId $user.Id -SoftwareType "teamsClient" -SoftwareVersion $tm.AvailableVersion -ErrorAction Stop
                                $sid = (Get-MgBetaTeamworkDeviceOperation -TeamworkDeviceId $user.Id -ErrorAction Stop | ForEach-Object {$_.Id} | Select-Object -Index '0')
                                # add data to did
                                $user.SID = $sid
                                $user.SoftwareType = 'TeamsClientUpdate'
                                $user.Status = 'Queued'
                                $user.AvailableVersion = $tm.AvailableVersion
                                $user.CurrentVersion = $tm.CurrentVersion
                                $user.SoftwareFreshness = $tm.SoftwareFreshness
                                $user.Sync = $tm.AvailableVersion
                                $targetid | Export-Csv "$path\tmp\did.csv" -NoTypeInformation
                            }
                        }
                }
        }
        catch {Write-Output ('No Dice: {0}' -f $user.DisplayName) ('Error: Update software - {0}' -f $_.Exception.Message) | Out-File "$path\logs\debug.txt" -Append}
        # re-check update software
        $targetid = Import-Csv "$path\tmp\did.csv"
        try {
            foreach ($user in $targetid | Where-Object {$_.HealthStatus -eq '' -and $_.Status -ne 'Successful'}) {
                Write-Progress -Activity ('Please wait...')
                $sw = Get-MgBetaTeamworkDeviceHealth -TeamworkDeviceId $user.Id -ErrorAction Stop
                $fw = $sw.softwareupdatehealth.FirmwareSoftwareUpdateStatus
                $cp = $sw.softwareupdatehealth.CompanyPortalSoftwareUpdateStatus
                $tm = $sw.softwareupdatehealth.TeamsClientSoftwareUpdateStatus
                if($user.SoftwareType -eq 'FirmwareUpdate' | Where-Object {$fw.SoftwareFreshness -ne 'latest'}) {
                    $user.HealthStatus = $uid.HealthStatus
                    $user.AvailableVersion = $fw.AvailableVersion
                    $user.CurrentVersion = $fw.CurrentVersion
                    $user.SoftwareFreshness = $fw.SoftwareFreshness
                    $targetid | Export-Csv "$path\tmp\did.csv" -NoTypeInformation
                }
                if($user.SoftwareType -eq 'CompanyPortalUpdate' | Where-Object {$cp.SoftwareFreshness -ne 'latest'}) {
                    $user.HealthStatus = $uid.HealthStatus
                    $user.AvailableVersion = $cp.AvailableVersion
                    $user.CurrentVersion = $cp.CurrentVersion
                    $user.SoftwareFreshness = $cp.SoftwareFreshness
                    $targetid | Export-Csv "$path\tmp\did.csv" -NoTypeInformation
                }
                if($user.SoftwareType -eq 'TeamsClientUpdate' | Where-Object {$tm.SoftwareFreshness -ne 'latest'}) {
                    $user.HealthStatus = $uid.HealthStatus
                    $user.AvailableVersion = $tm.AvailableVersion
                    $user.CurrentVersion = $tm.CurrentVersion
                    $user.SoftwareFreshness = $tm.SoftwareFreshness
                    $targetid | Export-Csv "$path\tmp\did.csv" -NoTypeInformation
                }
            }
        }
        catch {Write-Output ('No Dice: {0}' -f $user.DisplayName) ('Error: Re-check update software - {0}' -f $_.Exception.Message) | Out-File "$path\logs\debug.txt" -Append}
    $num = @()
    $targetusers = Import-Csv "$path\tmp\did.csv"
        foreach ($user in $targetusers) {
            $num += ($user.Id | Select-Object -Unique | Where-Object {$user.HealthStatus -eq 'nonUrgent'})
        }
    $counter = ($num | Select-Object -Unique).Count
    $displaycounter = '(' + $counter + ')'
    $flag = $true
    }
    catch {Write-Warning 'Please close the did.csv file and try again.'
    Pause
    }
    }
Write-Progress -Activity "Sleep" -Completed
Options_Menu_GUI
}

# Prompt the user to select a text file of one or more devices to be restarted
function Restart {
    if ($null -ne $batchfile -and $null -eq $flag2) {
        $confirm = Read-Host "Proceed with Teams Phone restart? [Y] Yes [N] No"
        Clear-Host
        if ($confirm -eq 'Y') {
            $targetid = Get-Content "$batchfile"
            $userarrayq = $targetid | Measure-Object
            $file = $userarrayq.count
            $i = 0
            try {
                foreach ($id in $targetid) {
                    $percentcomplete = ($i / $file) * 100; $i++
                    Write-Progress -Activity ('Restart on TeamworkDeviceId: ' + $id) -Status ('file ' + $i + ' of ' + $file) -PercentComplete $percentcomplete
                    $id = Get-MgBetaTeamworkDevice -TeamworkDeviceId $id -ErrorAction Stop | ForEach-Object {$_.Id}
                    $dn = Get-MgBetaTeamworkDevice -TeamworkDeviceId $id -ErrorAction Stop | ForEach-Object {$_.CurrentUser.DisplayName}
                    Restart-MgBetaTeamworkDevice -TeamworkDeviceId $id -ErrorAction Stop
                    $sid = (Get-MgBetaTeamworkDeviceOperation -TeamworkDeviceId $id -ErrorAction Stop | ForEach-Object {$_.Id} | Select-Object -Index '0')
                    # add data to res
                    $params = [ordered] @{
                        "Id" = $id
                        "DisplayName" = $dn
                        "Status" = 'Queued'
                        "SID" = $sid
                    }
                    $newrow = New-Object PSObject -Property $params
                    $importtable += $newrow | Export-Csv "$path\tmp\res.csv" -Append -NoTypeInformation
                }
            }
            catch {Write-Output ('No Dice: {0}' -f $id) ('Error: Device restart - {0}' -f $_.Exception.Message) | Out-File "$path\logs\debug.txt" -Append}
            $counter2 = $file
            $displaycounter2 = '(' + $counter2 + ')'
            $flag2 = $true
        }
    }elseif ($null -ne $batchfile -and $counter2 -gt 0 -and $null -ne $flag2) {
        $targetid = Import-Csv "$path\tmp\res.csv"
        $csv2 = "$path\tmp\res.csv"
        $ofile2 = New-Object System.IO.FileInfo $csv2
         try {
            # check if did.csv file is closed before continuing
            $ostream2 = $ofile2.Open([System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
            if ($ostream2) {$ostream2.Close()}
                try {
                    foreach ($user in $targetid) {
                        $ops = Get-MgBetaTeamworkDeviceOperation -TeamworkDeviceId $user.Id -ErrorAction Stop | Where-Object {$_.Id -eq $user.SID}
                        if ($ops.Status -eq 'Queued' -or $ops.Status -eq 'InProgress') {
                            Write-Progress ('Device restart ' + $ops.Status + ' on: ' + $user.DisplayName)
                            # add data to res
                            $user.Status = $ops.Status
                            $targetid | Export-Csv "$path\tmp\res.csv" -NoTypeInformation
                        }elseif ($ops.Status -eq 'Failed') {
                            Write-Progress -Activity ('Retrying Device restart on: ' + $user.DisplayName)
                            Restart-MgBetaTeamworkDevice -TeamworkDeviceId $user.Id -ErrorAction Stop
                            $sid = (Get-MgBetaTeamworkDeviceOperation -TeamworkDeviceId $user.Id -ErrorAction Stop | ForEach-Object {$_.Id} | Select-Object -Index '0')
                            # add data to res
                            $user.SID = $sid
                            $user.Status = 'Queued'
                            $targetid | Export-Csv "$path\tmp\res.csv" -NoTypeInformation
                        }elseif ($ops.Status -eq 'Successful') {
                            Write-Progress -Activity ('Please wait...')
                            # add data to res
                            $user.Status = $ops.Status
                            $user.SID = ''
                            $targetid | Export-Csv "$path\tmp\res.csv" -NoTypeInformation
                        }
                    }
                }
                catch {Write-Output ('No Dice: {0}' -f $user.DisplayName) ('Error: Verify restart - {0}' -f $_.Exception.Message) | Out-File "$path\logs\debug.txt" -Append}
                $num = @()
                $targetusers = Import-Csv "$path\tmp\res.csv"
                foreach ($user in $targetusers) {
                    $num += ($user.Id | Select-Object -Unique | Where-Object {$user.Status -ne 'Successful'})
                }
                $counter2 = ($num | Select-Object -Unique).Count
                $displaycounter2 = '(' + $counter2 + ')'
        }
        catch {Write-Warning 'Please close the res.csv file and try again.'
        Pause
        }
    }
Write-Progress -Activity "Sleep" -Completed
Options_Menu_GUI
}

# Remove session files on exit
function Quit {
    Stop-Transcript
    Disconnect-MgGraph
    Remove-Item "$path\tmp\*", "$path\logs\*"
}

#endregion

#region - SUB-FUNCTIONS

# User selects option A on the GUI which calls a function
function Option_1 {
Clear-Host
Archive
}

# User selects option B on the GUI which calls a function
function Option_2 {
Clear-Host
Clear_debug
}

# User selects option C on the GUI which calls a function
function Option_3 {
Clear-Host
Batch_file
}

# User selects option 1 on the GUI which calls a function
function Option_4 {
Clear-Host
Analyze_file
}

# User selects option 2 on the GUI which calls a function
function Option_5 {
Clear-Host
Update_software
}

# User selects option R on the GUI which calls a function
function Option_6 {
Clear-Host
Restart
}

# User selects option Q on the GUI which calls a function
function Option_Q {
Clear-Host
Quit
}

function Receive_Output_Red {
    process {Write-Host $_ -ForegroundColor Red}
}
function Receive_Output_Green {
    process {Write-Host $_ -ForegroundColor Green}
}

#endregion

#region - MENU GUI

function Options_Menu_GUI {
$customer = Get-MgOrganization | ForEach-Object DisplayName
do
{
Clear-Host
    Write-Host ''
    Write-Host "$customer $displaysite".ToUpper()
    Write-Host ''
    Write-Host ' Choose an option' -ForegroundColor Yellow
    Write-Host ''
    Write-Host ' 1.   Archive'
    Write-Host ''
    Write-Host ' 2.   Clear debug'
    Write-Host ''
    Write-Host " 3.   Select file $displayfilename"
    Write-Host ''
    Write-Host " 4.   Analyze file $ids"
    Write-Host ''
    if ($counter -gt 0 -and $flag -eq $false) {
    Write-Host ' 5.   Update software ' -NoNewLine
    Write-OutPut $displaycounter | Receive_Output_Red
    }elseif ($counter -gt 0 -and $flag -eq $true) {
    Write-Host ' 5.   Verify software ' -NoNewLine
    Write-OutPut $displaycounter | Receive_Output_Red
    }elseif ($counter -eq 0) {
    Write-Host ' 5.   Update software ' -NoNewLine
    Write-OutPut $displaycounter | Receive_Output_Green
    }else {
    Write-Host ' 5.   Update software'
    }
    Write-Host ''
    if ($counter2 -gt 0 -and $flag2 -eq $false) {
    Write-Host ' 6.   Device restart ' -NoNewLine
    Write-OutPut $displaycounter2 | Receive_Output_Red
    }elseif ($counter2 -gt 0 -and $flag2 -eq $true) {
    Write-Host ' 6.   Verify restart ' -NoNewLine
    Write-OutPut $displaycounter2 | Receive_Output_Red
    }elseif ($counter2 -eq 0) {
    Write-Host ' 6.   Device restart ' -NoNewLine
    Write-OutPut $displaycounter2 | Receive_Output_Green
    }else {
    Write-Host ' 6.   Device restart'
    }
    Write-Host ''
    Write-Host ' Q.   Quit'
    Write-Host ''
    $Menu = Read-Host -Prompt ' (1-6 or Q to Quit)'

    Switch ($Menu) {
            1
        {
            Option_1
        }
            2
        {
            Option_2
        }
            3
        {
            Option_3
        }
            4
        {
            Option_4
        }
            5
        {
            Option_5
        }
            6
        {
            Option_6
        }
            Q
        {
            Option_Q
            Clear-Host
            Exit
        }
    }
}
    until ($Menu -eq 'q')
}
Options_Menu_GUI

#endregion

