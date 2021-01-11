<#
    .SYNOPSIS
    Exports a OneNoteHierachy

    .NOTES
    - make sure to have the OneNote Client open
    - edit the config.cfg file with preferences
    - run this script
#>
(Get-ChildItem -Path "$PSScriptRoot\Modules" -Recurse -Filter '*.psm1' -Verbose).FullName | ForEach-Object { Import-Module $_ -Force }
(Get-ChildItem -Path "$PSScriptRoot\Export" -Recurse -Filter '*.psm1' -Verbose).FullName | ForEach-Object { Import-Module $_ -Force }

$config = Get-Config -path "$PSScriptRoot\Export\Config\export.cfg"
[array]$publishformats = Get-OneNotePublishFormats
$pandocIsNeeded = $false
$config.ExportFormat -split ',' -replace '^\s+|\s+$' | ForEach-Object {
        if (-not ($publishformats -contains $exportFormat)) {
            $pandocIsNeeded = $true
        }
}
if ($pandocIsNeeded -eq $true -and [System.Convert]::ToBoolean($Config.AutoDownloadPandoc) -eq $true) {
    $pandocExec = Get-PandocExecutable
    $config | Add-Member -Type NoteProperty -Name 'Pandoc' -Value $pandocExec -Force
}

Export-OneNoteHierarchy -Config $config

# TODO 1: https://github.com/PowerShell/vscode-powershell/issues/1856
#  After some time idle scope error appear, then need to run $error.Clear()
#  And it runs ok
# TODO 2: auto unfold pages based on the property since they apparently are not exported
# TODO 3: add a warning in the log if objects are encrypted
# todo validate that docx is specified if MD is chosen