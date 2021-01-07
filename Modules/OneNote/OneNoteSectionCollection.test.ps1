Import-Module "$PSScriptRoot\OneNoteHierarchy.psm1" -Force
Import-Module "$PSScriptRoot\OneNoteNotebookCollection.psm1" -Force
Import-Module "$PSScriptRoot\OneNoteNotebook.psm1" -Force

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
$sectionCollection = Get-OneNoteNotebookSectionCollection -Notebook $notebook
$tempId = $null
$sectionCollection | ForEach-Object { $tempId = $_.ID}

######################################################################################
# Get one section from section collection
######################################################################################
$section = Get-OneNoteSectionCollectionSection -SectionCollection $sectionCollection -ID $tempId
$section