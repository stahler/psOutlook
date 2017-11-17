# psOutlook
The psOutlook module has a few functions that will enable you to analyze outlook.  
Examples:
* Get-OutlookFolderCount: allows you to get counts of folders within a mailbox.  You can supply the root folder and if needed a list of specific subfolders you want a count.
* Get-OutlookSenders: Get a count of who is sending you emails
* Get-OutlookUnReadCount: Get a count of unread emails per folder

## Prerequisites
* PowerShell V3+
* Outlook - have not verified all versions, works with 2016.

## Getting started

## About the cmdlet

## Using this cmdlet

## Scenario 1: Looking at all folders in a mailbox
``` powershell
Get-OutlookFolderCount -root "Wes.Stahler@osumc.edu" | Out-GridView
```
## Scenario 2: Looking at specific folders in a mailbox
``` powershell
# define our list of folder that we want numbers from
$folders = '\\Spam Reporting (OSUWMC)\Inbox\Non Report SPAM\',
'\\Spam Reporting (OSUWMC)\Inbox\Report Phish Reports\By Date\',
'\\Spam Reporting (OSUWMC)\Inbox\Report Phish Reports\eusafe\',
'\\Spam Reporting (OSUWMC)\Inbox\Report Phish Reports\safe\'

Get-OutlookFolderCount -root "Spam Reporting (OSUWMC)" -folders $folders| Out-GridView
```

## Author
* Wes Stahler @stahler

## TODO
* Finish MD file
* Convert to module
* Add help
* Verify error handling
* Add more functions
  * Frequent Senders
  * Unread email per folder
  * Warning on count
