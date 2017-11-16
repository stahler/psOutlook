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

function Get-OutlookFolderCount {
    [cmdletbinding()]
    param(
        $root = "Spam Reporting (OSUWMC)",
        $folders = $null
    )

    process {
        $obj = @()
        try {
            $outlook = New-Object -ComObject Outlook.Application -ErrorAction Stop
            $RootFolder = $outlook.Session.Folders.Item($root).Folders
        }
        catch [System.Runtime.InteropServices.COMException] {
            Write-Warning "Outlook not installed or root folder does not exist"
            break
        }
        # iterates over all folders and ultimatly gets the full path and count
        foreach ($folder in $RootFolder) {
            foreach ($mailfolder in $folder.Folders) {
                $obj+=Get-MailboxFolder $mailfolder
            }
        }
        
        if ($folders) {
            $dt = Get-Date (Get-Date).AddDays(-1) -Format "MM-dd-yyyy"
            foreach ($folder in $folders) {
                $path = "$folder$dt"
                $obj.where({$_.path -eq $path})
            }
        } else {
            $obj
        }
    }
}

# define our list of folder that we want numbers from
$folders = '\\Spam Reporting (OSUWMC)\Inbox\Non Report SPAM\',
'\\Spam Reporting (OSUWMC)\Inbox\Report Phish Reports\By Date\',
'\\Spam Reporting (OSUWMC)\Inbox\Report Phish Reports\eusafe\',
'\\Spam Reporting (OSUWMC)\Inbox\Report Phish Reports\safe\'

#Get-OutlookFolderCount -root "Spam Reporting (OSUWMC)" | Out-GridView
#Get-OutlookFolderCount -root "Spam Reporting (OSUWMC))" -folders $folders | Out-GridView
#Get-OutlookFolderCount -root "Wes.Stahler@osumc.edu" | Out-GridView