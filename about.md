# About Software Updater for Microsoft Teams Phone

| [Home](README.md) | [About](about.md) | [User Guide](user.md) | [Support](support.md) | 
| --- | --- | --- | --- |

## Disclaimer
> [!IMPORTANT]
> _These samples are provided "as is" without warranty of any kind. SACOMS further disclaims all implied warranties including without limitation any implied warranties of merchantability or of fitness for a particular purpose. The entire risk arising out of the use or performance of the samples remains with you. In no event shall SACOMS be liable for any damages whatsoever (including, without limitation, damages for loss of business profits, business interruption, loss of business information, or other pecuniary loss) arising out of the use of or inability to use the samples._

## Overview
The Software Updater for Teams Phone is an alternative to using the Teams admin center (TAC) to manage software updates for Teams devices. This app uses PowerShell leveraging the Microsoft.Graph.Beta module to batch process multiple Teams devices requiring software updates. Additionally, it can also provide reporting with device metrics.

## Benefits
* Better autonomy for software updates.
* Batch processing of multiple devices at once.
* Increased efficiency
* Realtime progress metrics
* Works with all certified Teams phones.
  
## How it works
Instead of using the Teams admin center (TAC) via a web browser to schedule updates, this app uses Microsoft Graph PowerShell to connect to an M365 tenant.
Once connected, the 'Tenant Display Name' will change to that of the connected tenant, along with an intuitive gui menu.<br/>

<img src="https://github.com/SimonJonesZEE/MicrosoftTeamsPhone-SoftwareUpdater/blob/main/assets/options-menu-gui.png" width="350" height="300"> 

There are multiple options to choose from.
#### 1. Archive
> Creates a compressed ZIP archive of the session files.
#### 2. Clear debug
> Clears the contents of the debug file.
#### 3. Select file 
> This will allow you to browse for a file using file explorer.
#### 4. Analyze file 
> This will analyze the selected file which contains one or more TeamworkDeviceIds.
#### 5. Update software
> Schedules software updates on Teams Phone devices.
#### 6. Device restart 
> Initiates a restart of all TeamworkDeviceId's within the selected text file.
#### Q. Quit
> This will close and exit the PowerShell GUI.
   
## Page info

| Page | About |
| :--- | :--- |
| Author | Simon Jones ([@simonjoneszee](https://github.com/simonjoneszee)) |
| **Version** | 1.0.1 |
| **Date** | 02/10/2024 |
