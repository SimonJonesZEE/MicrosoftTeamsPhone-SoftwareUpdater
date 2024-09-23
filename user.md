# User Guide

| [Home](README.md) | [About](about.md) | [User Guide](user.md) | [Support](support.md) | 
| --- | --- | --- | --- |

## 1. Pre-requisites
* PowerShell modules:
  * Microsoft.Graph.Beta (latest)
* Microsoft Entra roles
    - Cloud Application Administrator (preferred)
    - Teams Admnistrator
 > [!NOTE]
  >  The modules folder structure will be created for you when its first run.

## 2. Installation
* Install the script module to your **$HOME\Documents\WindowsPowerShell\Modules** directory.
> [!WARNING]
  >  Please do not install this module to the **%Windir%\System32\WindowsPowerShell\v1.0\Modules** directory as this is reserved for modules that ship with Windows.

## 3. Execution
1. Launch the PowerShell console and type **Import-Module TeamsPhoneSoftwareUpdater**
2. This will prompt for you to sign in to your account using MFA.<br/>

   <img src="https://github.com/SimonJonesZEE/MicrosoftTeamsPhone-SoftwareUpdater/blob/main/assets/sign-in.png" width="350" height="250"> 
   
3. Once signed in, you will see a GUI options menu, and the 'TENANT DISPLAY NAME' will change to that of the connected tenant.<br/>

   <img src="https://github.com/SimonJonesZEE/MicrosoftTeamsPhone-SoftwareUpdater/blob/main/assets/options-menu-gui.png" width="350" height="290"> 
   
## 4. Options

* **1. Archive** <br/> This will create a compressed ZIP archive from specified files in the 'tmp' and 'logs' directories and places the ZIP file in the 'arc' directory.
> [!TIP]
> This is useful if you want to archive the session files before exiting or selecting a new file.

* **2. Clear debug** <br/> This will clear the contents of the 'debug.txt' located in the 'logs' directory.
> [!NOTE]
> The debug.txt file is only generated if an exception occurs.
  
* **3. Select file** <br/> This will default to the 'data' directory when selecting a text file; alternatively, you can browse elsewhere using the file explorer.

* **4. Analyze file** <br/> This will perform an analysis of the selected text file, which contains one or more TeamworkDeviceIds.

* **5. Update software** <br/> This will schedule software updates on devices if applicable.
> [!NOTE]
> This option will only work if the file analysis detects one or more devices requiring updates.

* **6. Restart** <br/> This will initiate a restart of all TeamworkDeviceIds within the selected text file.

* **7. Quit** <br/> This will close and exit the PowerShell GUI removing any session files contained within the 'tmp' and 'logs' directories.
> [!TIP]
> Remember to archive before exiting the application if you want an audit trail of the previous updates.

## 5. How-to
**Step 1:** Once authenticated to start using the Microsoft Teams Phone Software Updater, we must first select a text file that contains one or more TeamworkDeviceIds by selecting option **3. Select file**.<br/>

You will then be prompted for user input.<br/>
   
         Are you sure you want to load a new file? [Y] Yes [N] No: Y
   
> [!NOTE]
> Inputting [N] No or left blank will return you back to the main menu.
   
        Enter a site name or code: LDNOffice
        
> [!NOTE]
> Please omit the use of spaces in the site name or code.<br/>

Using the file explorer, find and select the text file you want to use.<br/>
   
<img src="https://github.com/SimonJonesZEE/MicrosoftTeamsPhone-SoftwareUpdater/blob/main/assets/notepad.png" width="350" height="250"> 

Then **Press [C] to continue:**

The site name or code (LDNOFFICE) and the selected file (devices.txt) will appear as a visual and functional representation of the current session.<br/>

<img src="https://github.com/SimonJonesZEE/MicrosoftTeamsPhone-SoftwareUpdater/blob/main/assets/selected-file.png" width="350" height="290"> 
   
**Step 2:** The next thing is to analyse the selected text file by using option **4. Analyze file**, this will verify each TeamworkDeviceId to see if there any updates required.<br/>

The main updates this app verifies are:<br/>
- [ ] Teams admin agent
- [x] Firmware
- [x] Company portal app
- [ ] Oem agent app
- [x] Teams app

After the analysis has completed, you will see how many devices have been verified next to **4. Analyze file (3)** and how many devices require updates next to **5. Update software (2)**.

<img src="https://github.com/SimonJonesZEE/MicrosoftTeamsPhone-SoftwareUpdater/blob/main/assets/device-update.png" width="350" height="290"> 

Every time a file analysis is performed, a 'did.csv' session file is created, which is located in the 'tmp' directory.<br/>

<img src="https://github.com/SimonJonesZEE/MicrosoftTeamsPhone-SoftwareUpdater/blob/main/assets/file-analysis.png" width="1200" height="190">

> **HealthStatus:**
This will be stamped with 'nonUrgent' if the respective software type of the device requires an update.

> [!NOTE]
> If the HealthStatus shows as being 'offline' against a subset of device software types, they will be excluded from being updated.

> **Id:**
The identification of the Teams Phone device.

> **DisplayName:**
The display name of the Teams Phone device.

> **SoftwareType:**
This is duplicated three times for each TeamworkDeviceId, one for each software type being verified.

> **Status:**
Provides an update on the progress of the software update, e.g., 'Queued', 'InProgress, 'Successful'.

> **SID:**
This is a unique identifier that is generated automatically for each update.

> **AvailableVersion:**
Only populated if there is an available update for the device.

> **CurrentVersion:**
The current version of software the device is running.

> **SoftwareFreshness:**
This will either state 'updateAvailable' or 'latest'.

> **Rep:**
This is used to record the update being applied to the device and compared against the device's current version on the tenant. If the equality comparison in the script matches these two values, the replication in the tenant is complete, and the 'Rep' value is then null. If the equality comparison does not match, then the Rep value will be persistent until there is a match, preventing accidental re-queuing of updates and enabling the script to check for further step updates against the device's current version of the software once replication is complete.

☑️**Step 3:** After file analysis please review the 'did.csv' session file, and then select option **5. Update software (2)**.


## Page info

| Page | User Guide |
| :--- | :--- |
| Author | Simon Jones ([@simonjoneszee](https://github.com/simonjoneszee)) |
| **Version** | 1.0 |
| **Date** | 13/09/2024 |
