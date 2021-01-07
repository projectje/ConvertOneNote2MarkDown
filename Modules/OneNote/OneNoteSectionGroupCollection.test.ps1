Import-Module "$PSScriptRoot\OneNoteHierarchy.psm1" -Force
Import-Module "$PSScriptRoot\OneNoteNotebookCollection.psm1" -Force
Import-Module "$PSScriptRoot\OneNoteNotebook.psm1" -Force
Import-Module "$PSScriptRoot\OneNoteSection.psm1" -Force
Import-Module "$PSScriptRoot\OneNoteSectionGroup.psm1" -Force
Import-Module "$PSScriptRoot\OneNoteSectionGroupCollection.psm1" -Force

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
Write-Host "Sectiongroup collection items which have sectiongroups :" -BackgroundColor Green -ForegroundColor Black
$tempIdForSectionGroup = $null
$sectionGroupCollectionInRoot = Get-OneNoteSectionGroupCollectionSectionGroupCollection -SectionGroup $sectiongroupRoot
$sectionGroupCollectionInRoot

######################################################################################
# Get-OneNoteSectionGroupCollectionSectionCollection
######################################################################################
# just pick one with a section
Write-Host "Sectioncollection items which have section :" -BackgroundColor Green -ForegroundColor Black
$tempIdForSection = $null
$sectionCollectionInRoot = Get-OneNoteSectionGroupCollectionSectionCollection -SectionGroup $sectiongroupRoot
$sectionCollectionInRoot