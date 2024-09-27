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

* **1. Archive** <br/> This will create a compressed ZIP archive from specified files in the `tmp` and `logs` directories and places the ZIP file in the `arc` directory.
> [!TIP]
> This is useful if you want to archive the session files before exiting or selecting a new file.

* **2. Clear debug** <br/> This will clear the contents of the `debug.txt` located in the `logs` directory.
> [!NOTE]
> The debug.txt file is only generated if an exception occurs.
  
* **3. Select file** <br/> This will default to the `data` directory when selecting a text file; alternatively, you can browse elsewhere using the file explorer.

* **4. Analyze file** <br/> This will perform an analysis of the selected text file, which contains one or more TeamworkDeviceIds.

* **5. Update software** <br/> This will schedule software updates on devices if applicable.
> [!NOTE]
> This option will only work if the file analysis detects one or more devices requiring updates.

* **6. Restart** <br/> This will initiate a restart of all TeamworkDeviceIds within the selected text file.

* **Q. Quit** <br/> This will close and exit the PowerShell GUI removing any session files contained within the `tmp` and `logs` directories.
> [!TIP]
> Remember to archive before exiting the application if you want an audit trail of the previous updates.

## 5. How-to
<mark>**STEP 1:**</mark>  

Once authenticated to start using the Microsoft Teams Phone Software Updater, we must first select a text file that contains one or more TeamworkDeviceIds by selecting option **3. Select file**.  

You will then be prompted for user input.  
   
         Are you sure you want to load a new file? [Y] Yes [N] No: Y
   
> [!NOTE]
> Inputting [N] No or left blank will return you back to the main menu.
   
        Enter a site name or code: LDNOffice
        
> [!NOTE]
> Please omit the use of spaces in the site name or code.  

Using the file explorer, find and select the text file you want to use.  
   
<img src="https://github.com/SimonJonesZEE/MicrosoftTeamsPhone-SoftwareUpdater/blob/main/assets/notepad.png" width="350" height="250"> 

Then **Press [C] to continue:**

The site name or code '(LDNOFFICE)' and the selected file '(devices.txt)' will appear as a visual and functional representation of the current session.  

<img src="https://github.com/SimonJonesZEE/MicrosoftTeamsPhone-SoftwareUpdater/blob/main/assets/selected-file.png" width="350" height="290"> 
   
<mark>**STEP 2:**</mark>  

The next thing is to analyse the selected text file by using option **4. Analyze file**, this will verify each TeamworkDeviceId to see if there any updates required.  

The main updates this app verifies are:  
- [ ] Teams admin agent
- [x] Firmware
- [x] Company portal app
- [ ] Oem agent app
- [x] Teams app

Once the analysis has completed, you will see how many devices have been verified next to **4. Analyze file** and how many devices require updates next to **5. Update software**.  

<img src="https://github.com/SimonJonesZEE/MicrosoftTeamsPhone-SoftwareUpdater/blob/main/assets/device-update.png" width="350" height="290"> 

This will also generate a `did.csv` file, which is located in the `tmp` directory.  

<img src="https://github.com/SimonJonesZEE/MicrosoftTeamsPhone-SoftwareUpdater/blob/main/assets/file-analysis.png">

> **HealthStatus:**
This will be stamped with `nonUrgent` if the respective software type requires an update.

> [!NOTE]
> If the HealthStatus shows as being `offline` against a subset of device software types, they will be excluded from being updated.

> **Id:**
The identification of the Teams Phone device.

> **DisplayName:**
The display name of the Teams Phone device.

> **SoftwareType:**
This is duplicated three times for each TeamworkDeviceId, one for each software type being verified.

> **Status:**
Provides an update on the progress of the software update, e.g., `Queued`, `InProgress`, `Successful`, `Failed`

> **SID:**
This is a unique identifier that is generated automatically for each update.

> **AvailableVersion:**
Only populated if there is an available update for the device.

> **CurrentVersion:**
The current version of software the device is running.

> **SoftwareFreshness:**
This will either state `updateAvailable` or `latest`.

> **Sync:**
This is used to record the update being applied to the device and compared against the device's current version on the tenant. If the equality comparison in the script matches these two values, synchronisation in the tenant is complete, and the **Sync** value is null. If the equality comparison does not match, the **Sync** value will be persistent until there is a match, preventing accidental re-queuing of updates and enabling the script to check for further step updates against the device's current version of the software once synchronisation is complete.

<mark>**STEP 3:**</mark>  

Select option **5. Update software** to kickstart the process; the option will dynamically change to **5. Verify software** once completed.

<img src="https://github.com/SimonJonesZEE/MicrosoftTeamsPhone-SoftwareUpdater/blob/main/assets/device-update2.png" width="350" height="290"> 

If you open the `did.csv` file, you will notice some changes. 

<img src="https://github.com/SimonJonesZEE/MicrosoftTeamsPhone-SoftwareUpdater/blob/main/assets/file-analysis2.png">  

The **Status** field is showing the update as `Queued` and the **SID** field contains the unique id of the update, and the **Sync** field contains a snapshot of the available update being applied.  

> [!TIP]
> You can repeatedly run option **5. Verify software**, which will update the `did.csv` file in order for you to track the progress of the updates being applied to the devices and for the execution of further step updates if required.

<img src="https://github.com/SimonJonesZEE/MicrosoftTeamsPhone-SoftwareUpdater/blob/main/assets/file-analysis3.png"> 

When the updates have been applied successfully and dependent on the update being applied, the **Sync** field will be blank, which indicates synchronisation of the update in Teams is complete, and the **Status** field records the update as `Successful` and all other fields are refreshed.

<img src="https://github.com/SimonJonesZEE/MicrosoftTeamsPhone-SoftwareUpdater/blob/main/assets/file-analysis4.png">  

> [!TIP]
> If the **Status** shows as `Failed` against a device update, then select option **5. Verify software** to retry the update on the device(s).

> [!IMPORTANT]
> In some instances, the **HealthStatus** of a device may change and show as `offline` which is common when updating firmware; when this happens, manually reboot the device or if the device is accessible over the network log into the GUI and reboot the device.

> [!WARNING]
> Updating device firmware will take longer to synchronize than a Company Portal or a Teams Client update, so periodically re-select option **5. Verify software** to check it's progress allowing enough time inbetween whilst the device no doubt reboots whilst processing this change.

With all things being equal, option **5. Verify software** will have reverted back to **5. Update software**, indicating '(0)' confirming all devices are running on the latest software.  

<img src="https://github.com/SimonJonesZEE/MicrosoftTeamsPhone-SoftwareUpdater/blob/main/assets/device-update3.png" width="350" height="290"> 

At this point you can either select option **3. Select file** to update different devices on the tenant or you can select option **Q. Quit** and re-run this application to update devices on another tenant.

<mark>**STEP 4: (optional)**</mark>  

If you want to schedule a restart of one or more Teams devices then select option **6. Device restart**.

> [!TIP]
> File analysis is not required for this option.

<img src="https://github.com/SimonJonesZEE/MicrosoftTeamsPhone-SoftwareUpdater/blob/main/assets/device-restart.png" width="350" height="290"> 

You will then be prompted for user input.  
   
         Proceed with Teams Phone restart? [Y] Yes [N] No: Y  
         
> [!NOTE]
> Inputting [N] No or left blank will return you back to the main menu.

> [!WARNING]
> This will reboot device(s) and will take them out of service momentarily whilst they reboot.  
> Please proceed with caution when using this option.

When option **6. Device restart** is initiated, the option will dynamically change to **5. Verify restart (1)**, indicating how many devices are scheduled for a reboot. 

<img src="https://github.com/SimonJonesZEE/MicrosoftTeamsPhone-SoftwareUpdater/blob/main/assets/device-restart2.png" width="350" height="290"> 

Like before you can repeatedly run option **6. Verify restart**, which will update the `res.csv` file located in the `tmp` directory in order for you to track the progress of the device restarts.

You're all caught up. 😊  

## Page info

| Page | User Guide |
| :--- | :--- |
| Author | Simon Jones ([@simonjoneszee](https://github.com/simonjoneszee)) |
| **Version** | 1.0 |
| **Date** | 26/09/2024 |
