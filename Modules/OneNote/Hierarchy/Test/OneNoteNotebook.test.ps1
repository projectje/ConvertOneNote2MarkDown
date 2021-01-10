Import-Module "$PSScriptRoot\..\OneNoteHierarchy.psm1" -Force
Import-Module "$PSScriptRoot\..\OneNoteNotebookCollection.psm1" -Force
Import-Module "$PSScriptRoot\..\OneNoteNotebook.psm1" -Force

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

######################################################################################
# Get-OneNoteNoteBook
# Dump content
######################################################################################
Write-Host "Content of notebook 1:" -BackgroundColor Green -ForegroundColor Black
$notebook = Get-OneNoteNotebook -NotebookCollection $notebookCollection -NotebookItem $item
$notebook

######################################################################################
# Get-OneNoteNoteBookSectionCollection
# Dump sections
######################################################################################
Write-Host "Sections of notebook 0:" -BackgroundColor Green -ForegroundColor Black
$sectionCollection = Get-OneNoteNotebookSectionCollection -Notebook $notebook
$sectionCollection

######################################################################################
# Get-OneNoteNoteBookSectionGroupCollection
# Dump section groups
######################################################################################
Write-Host "Sectiongroups of notebook 0:" -BackgroundColor Green -ForegroundColor Black
$sectionGroupCollection = Get-OneNoteNotebookSectionGroupCollection -Notebook $notebook -Include_Recyclebin $false
$sectionGroupCollection
