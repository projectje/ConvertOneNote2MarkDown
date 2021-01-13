<#
    .SYNOPSIS
    Exports a OneNoteHierachy

    .NOTES
    - make sure to have the OneNote Client open
    - edit the config.cfg file with preferences
#>
(Get-ChildItem -Path "$PSScriptRoot\Modules" -Recurse -Filter '*.psm1' -Verbose).FullName | ForEach-Object { Import-Module $_ -Force }
(Get-ChildItem -Path "$PSScriptRoot\Export" -Recurse -Filter '*.psm1' -Verbose).FullName | ForEach-Object { Import-Module $_ -Force }
Export-OneNoteHierarchy -Config (Export-Config -Path "$PSScriptRoot\Export\Config\export.cfg")

# TODO: auto unfold pages based on the property since they apparently are not exported
# TODO: add a warning in the log if objects are encrypted
# todo validate that docx is specified if MD is chosen
#$global:error
$Error.Clear()