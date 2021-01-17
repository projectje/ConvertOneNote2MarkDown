<#
    .SYNOPSIS
    Exports a OneNoteHierachy

    .NOTES
    - make sure to have the OneNote Client open
    - edit the config.cfg file with preferences
#>
(Get-ChildItem -Path "$PSScriptRoot\Modules" -Recurse -Filter '*.psm1' -Verbose).FullName | ForEach-Object { Import-Module $_ -Force }
(Get-ChildItem -Path "$PSScriptRoot\ScriptModules" -Recurse -Filter '*.psm1' -Verbose).FullName | ForEach-Object { Import-Module $_ -Force }
(Get-ChildItem -Path "$PSScriptRoot\Config" -Recurse -Filter '*.psm1' -Verbose).FullName | ForEach-Object { Import-Module $_ -Force }

$config = Get-OneNotePublishConfiguration -Path "$PSScriptRoot\Config\publish.cfg"
Publish-OneNoteHierarchyPages -Config $config