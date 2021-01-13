(Get-ChildItem -Path "$PSScriptRoot\..\..\Modules" -Recurse -Filter '*.psm1' -Verbose).FullName | ForEach-Object { Import-Module $_ -Force }
(Get-ChildItem -Path "$PSScriptRoot\..\" -Recurse -Filter '*.psm1' -Verbose).FullName | ForEach-Object { Import-Module $_ -Force }

$config = Export-Config -Path "$PSScriptRoot\..\Config\export.cfg"
$config

