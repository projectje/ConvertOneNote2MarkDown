Import-Module "$PSScriptRoot\..\OneNoteHierarchy.psm1" -Force
Import-Module "$PSScriptRoot\..\OneNoteNotebookCollection.psm1" -Force
Import-Module "$PSScriptRoot\..\OneNoteNotebook.psm1" -Force
Import-Module "$PSScriptRoot\..\OneNoteSection.psm1" -Force
Import-Module "$PSScriptRoot\..\OneNoteSectionGroup.psm1" -Force
Import-Module "$PSScriptRoot\..\OneNoteSectionGroupCollection.psm1" -Force

######################################################################################
# Setup (tested in other tests)
######################################################################################
$hierarchy = Get-OneNoteHierarchy
$notebookCollection = Get-OneNoteNotebookCollection  -Hierarchy $hierarchy
$count = Get-OneNoteNotebookCollectionCount -NotebookCollection $notebookCollection
if ($count -lt 1) {
    Write-host "Collection contains no notebooks"
    Exit;
}
$item = 1
$notebook = Get-OneNoteNotebook -NotebookCollection $notebookCollection -NotebookItem $item
$sectiongroupRoot = Get-OneNoteNotebookSectionGroupCollection -Notebook $notebook -Include_Recyclebin $false
######################################################################################
Write-Host "All entries in root sectiongroup :" -BackgroundColor Green -ForegroundColor Black
$sectiongroupRoot

######################################################################################
# Get-OneNoteSectionGroupCollectionSectionGroupCollection
######################################################################################
# just pick one with a sectiongroup
Write-Host "Sectiongroup collection in root sectiongroup :" -BackgroundColor Green -ForegroundColor Black
$tempIdForSectionGroup = $null
$sectionGroupCollectionInRoot = Get-OneNoteSectionGroupCollectionSectionGroupCollection -SectionGroup $sectiongroupRoot
$sectionGroupCollectionInRoot
$sectionGroupCollectionInRoot | ForEach-Object {
    $tempIdForSectionGroup = $_.ID
}

######################################################################################
# Get-OneNoteSectionGroupCollectionSectionCollection
######################################################################################
# just pick one with a section
Write-Host "Sectioncollection in root sectiongroup :" -BackgroundColor Green -ForegroundColor Black
$tempIdForSection = $null
$sectionCollectionInRoot = Get-OneNoteSectionGroupCollectionSectionCollection -SectionGroup $sectiongroupRoot
$sectionCollectionInRoot
$sectionCollectionInRoot | ForEach-Object {
    $tempIdForSection = $_.ID
}

######################################################################################
# Get-OneNoteSectionGroupSectionGroupCollection
######################################################################################
$sectiongroupRoot.ID

Write-Host "One Sectioncollection in root sectiongroup with id " $tempIdForSection ":" -BackgroundColor Green -ForegroundColor Black
$sectionCollection = Get-OneNoteSectionGroupSectionGroupCollection -SectionGroup $sectiongroupRoot -ID $tempIdForSection
$sectionCollection

######################################################################################
# Get-OneNoteSectionGroupSectionCollection
######################################################################################
Write-Host "One Sectiongroupcollection in root sectiongroup with id " $tempIdForSectionGroup ":" -BackgroundColor Green -ForegroundColor Black
$sectiongroup = Get-OneNoteSectionGroupSectionCollection -SectionGroup $sectiongroupRoot -ID $tempIdForSectionGroup
$sectiongroup
