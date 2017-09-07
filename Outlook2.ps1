function Get-MailboxFolder {
    param($folder)

    [pscustomobject]@{
        Path = $folder.FullFolderPath
        Count = $folder.items.count
    }

    foreach ($f in $folder.folders) {
        Get-MailboxFolder $f
    }
}

$obj = @()

$outlook = New-Object -ComObject Outlook.Application
$RootFolder = $outlook.Session.Folders.Item("Spam Reporting (OSUWMC)").Folders
foreach ($folder in $RootFolder) {
    foreach ($mailfolder in $folder.Folders) {
        $obj+=Get-MailboxFolder $mailfolder
    }
}

# format the date to match the Outlook folder format
$dt = Get-Date (Get-Date).AddDays(-1) -Format "MM-dd-yyyy"

# define our list of folder that we want numbers from
$folders = '\\Spam Reporting (OSUWMC)\Inbox\Non Report SPAM\',
           '\\Spam Reporting (OSUWMC)\Inbox\Report Phish Reports\By Date\',
           '\\Spam Reporting (OSUWMC)\Inbox\Report Phish Reports\eusafe\',
           '\\Spam Reporting (OSUWMC)\Inbox\Report Phish Reports\safe\'

foreach ($folder in $folders) {
    $path = "$folder$dt"
    $obj.where({$_.path -eq $path})
}

