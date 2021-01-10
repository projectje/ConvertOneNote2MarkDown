Import-Module "$PSScriptRoot\..\OneNoteHierarchy.psm1" -Force

#################################################################################
# Get-OneNoteHierarchy
# with default parameters (none)
$hierarchy = Get-OneNoteHierarchy

#
# Array of objects is returned (context and notebooks)
#
Write-Host "Array of objects is returned:" -BackgroundColor Blue -ForegroundColor White
$type = $hierarchy.GetType()
if ($type.BaseType -eq [System.Array]) {
    Write-Host "OK"
}

#
# GetMembers, should contain "Notebooks"
#
Write-Host "Object should contain member Notebooks:" -BackgroundColor Blue -ForegroundColor White
foreach ($property in $hierarchy | get-member) {
    if ($property.MemberType -eq "Property" -and $property.Name -eq "Notebooks") {
        Write-Host "OK"
    }
}

#
# Get notebooks in the XML
#
Write-Host "Content of notebooks:" -BackgroundColor Blue -ForegroundColor White
$notebooks = $hierarchy.Notebooks
if ($null -ne $notebooks) {
    Write-Host "OK"
}

#
# Dump of contents of notebooks
#
Write-Host "Content of notebooks:" -BackgroundColor Green -ForegroundColor Black
$hierarchy.Notebooks