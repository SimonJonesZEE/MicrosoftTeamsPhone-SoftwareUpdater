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

## 2. How-to
1. Launch the PowerShell console and type **Import-Module TeamsPhoneSoftwareUpdater**
2. This will prompt for you to sign in to your account using MFA.<br/>

   <img src="https://github.com/SimonJonesZEE/MicrosoftTeamsPhone-SoftwareUpdater/blob/main/assets/sign-in.png" width="350" height="250"> 
   
3. Once signed in, you will see a GUI options menu, and the 'TENANT DISPLAY NAME' will change to that of the connected tenant.<br/>

   <img src="https://github.com/SimonJonesZEE/MicrosoftTeamsPhone-SoftwareUpdater/blob/main/assets/options-menu-gui.png" width="350" height="300"> 
   
## 3. Options

* **Option 1. Archive** <br/> This will create a compressed ZIP archive from specified files in the 'tmp' and 'logs' directories and places the ZIP file in the 'arc' directory.
> [!TIP]
> This is useful if you want to archive the session files before exiting or selecting a new file.
* **Option 2. Clear debug** <br/> This will clear the contents of the 'debug.txt' if generated; this file is located in the 'logs' directory.


## Page info

| Page | User Guide |
| :--- | :--- |
| Author | Simon Jones ([@simonjoneszee](https://github.com/simonjoneszee)) |
| **Version** | 1.0 |
| **Date** | 13/09/2024 |
