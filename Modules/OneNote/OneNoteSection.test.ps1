Import-Module "$PSScriptRoot\OneNoteHierarchy.psm1" -Force
Import-Module "$PSScriptRoot\OneNoteNotebookCollection.psm1" -Force
Import-Module "$PSScriptRoot\OneNoteNotebook.psm1" -Force
Import-Module "$PSScriptRoot\OneNoteSectionCollection.psm1" -Force
Import-Module "$PSScriptRoot\OneNoteSection.psm1" -Force

######################################################################################
# Setup (tested in other tests)
######################################################################################
$hierarchy = Get-OneNoteHierarchy
$notebookCollection = Get-OneNoteNotebookCollection  -Hierarchy $hierarchy
$count = Get-OneNoteNoteBookCollectionCount -NotebookCollection $notebookCollection
if ($count -lt 1) {
    Write-host "Collection contains no notebooks"
    Exit;
}
$item = 1
$notebook = Get-OneNoteNotebook -NotebookCollection $notebookCollection -NotebookItem $item
$sectionCollection = Get-OneNoteNotebookSectionCollection -Notebook $notebook
$testId = $null
$sectionCollection | ForEach-Object { $testId = $_.ID}
Write-Host "Section {$testId}:" -BackgroundColor Green -ForegroundColor Black
$section = Get-OneNoteSectionCollectionSection -SectionCollection $sectionCollection -ID $testId

######################################################################################
# Get page collection in one section
######################################################################################
Write-Host "Get Pages for Section {$testId}:" -BackgroundColor Green -ForegroundColor Black
$pageCollection = Get-OneNoteSectionPageCollection -Section $section
$pageCollection

######################################################################################
# Get amount of pages in pagecollection in one section
######################################################################################
$pagecount = Get-OneNoteSectionPageCount -Section $section
$pagecount