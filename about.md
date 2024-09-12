# About Microsoft Teams Phone Software Updater

| [Home](README.md) | [About](about.md) | [Considerations](considerations.md) | [Deployment Guide](deployment.md) | [Support](support.md) | 
| --- | --- | --- | --- | --- |

## Disclaimer
> [!IMPORTANT]
> These samples are provided "as is" without warranty of any kind. SACOMS further disclaims all implied warranties including without limitation any implied warranties of merchantability or of fitness for a particular purpose. The entire risk arising out of the use or performance of the samples remains with you. In no event shall SACOMS be liable for any damages whatsoever (including, without limitation, damages for loss of business profits, business interruption, loss of business information, or other pecuniary loss) arising out of the use of or inability to use the samples.

## Overview
The Microsoft Teams Phone Software Updater is an alternative to using the Teams admin center (TAC) to manage software updates for Teams devices. This app uses PowerShell leveraging the Microsoft.Graph.Beta module to batch process multiple Teams devices requiring software updates. Additionally, it can also provide reporting with device metrics.

## Benefits
* Better autonomy for software updates.
* Batch processing of multiple devices at once.
* Increased efficiency
* Realtime progress metrics
* Works with all certified Teams phones.
  
## How it works
Instead of assigning a phone number to every user, you use the phone number of a resource account associated with an Auto attendant for inbound and outbound PSTN calls. The Shared Calling policy configures the resource account used for outbound calls and emergency numbers that are used as emergency callback numbers.

![image](https://github.com/SimonJonesZEE/MicrosoftTeamsPhone-SoftwareUpdater/blob/main/assets/options-menu-gui.png)

## Page info

| Page | About |
| :--- | :--- |
| Author | Simon Jones ([@simonjonesZEE](https://github.com/simonjonesZEE)) |
| **Version** | 1.0 |
| **Date** | 12/09/2024 |
