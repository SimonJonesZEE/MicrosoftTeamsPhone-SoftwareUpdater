Microsoft Teams Phone Software Updater
This tool will allow you to batch one or more TeamworkDeviceId's into a text file, the script
will analyse the list of TeamworkDeviceId's and determine which device needs upgrading and schedule the upgrade.
All this is done via an intuitive PowerShell GUI, there are 7 x options to choose from...

* A - Archive: This option will archive session data, archiving the 'Query' and 'Status' reports located in the '.\arc' folder.
* B - Batch File (): This option will require the user to select a text file which contains one or more TeamworkDeviceId's located in the '.\data' folder. When a text file has been selected the filename will appear within the () to signify which batch file is being worked on.
* C - Clear Debugs: This option will clean the debugs.txt file if present, this file will be located in the '.\logs' folder.
* 1 - Query Devices: This option will perform a query based on the selected 'Batch File' which contains the TeamworkDeviceId(s). After the query has completed an Query.csv will appear in the '\exp' folder which the user can then analyse.
* 2 - Device Updates (): This option will only display a number within the brackets () of devices requiring an update after option '1' has completed.




