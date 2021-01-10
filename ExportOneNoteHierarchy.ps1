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

Export-OneNoteHierarchy -Config (Get-Config -path "$PSScriptRoot\Export\Config\export.cfg")

# TODO 1: a specific section 