# Exploring Outlook with PowerShell
$outlook = New-Object -ComObject Outlook.Application
$RootFolder = $outlook.Session.Folders | Where-Object FolderPath -eq '\\Spam Reporting (OSUWMC)'
foreach ($folder in $RootFolder) {
    foreach ($mailfolder in $folder.Folders) {
        foreach ($subMailFolder in $mailfolder.Folders) {
            $cnt = $submailfolder.UnReadItemCount
            if ($cnt -gt 0) {
                "{0}`t{1}" -f $submailfolder.UnReadItemCount, $submailfolder.FolderPath
            }
        }
    }
}