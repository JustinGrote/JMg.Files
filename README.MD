# JMg.Files Module

This module makes working with Files (aka DriveItems) easier with Microsoft Graph. The inspiration was the fact that
PnP.Powershell as of May 2022 did not work with the Sites.Selected App Scope that allows granular site collection access.

## Quickstart

```powershell
Install-Module JMg.Files

#Gets all files at the root of your OneDrive
$files = Get-JMgDriveChildItem

#Downloads the files to temp
$files | Save-JMgDriveItem

#Upload Files to the root of your onedrive
Get-ChildItem -File /to/my/path | Push-JMgDriveItem

```

See the help on each command for more examples.
