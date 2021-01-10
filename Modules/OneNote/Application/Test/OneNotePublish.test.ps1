Import-Module "$PSScriptRoot\..\..\Hierarchy\OneNoteHierarchy.psm1" -Force
Import-Module "$PSScriptRoot\..\..\Hierarchy\OneNoteNotebookCollection.psm1" -Force
Import-Module "$PSScriptRoot\..\..\Hierarchy\OneNoteNotebook.psm1" -Force
Import-Module "$PSScriptRoot\..\..\Hierarchy\OneNoteSectionCollection.psm1" -Force
Import-Module "$PSScriptRoot\..\..\Hierarchy\OneNoteSection.psm1" -Force
Import-Module "$PSScriptRoot\..\..\Hierarchy\OneNotePageCollection.psm1" -Force
Import-Module "$PSScriptRoot\..\..\Hierarchy\OneNotePage.psm1" -Force
Import-Module "$PSScriptRoot\..\OneNotePublish.psm1" -Force

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
$item = 0
$notebook = Get-OneNoteNotebook -NotebookCollection $notebookCollection -NotebookItem $item
$sectionCollection = Get-OneNoteNotebookSectionCollection -Notebook $notebook
$testId = $null
$sectionCollection | ForEach-Object { $testId = $_.ID}
#$sectionCollection
Write-Host "Section {$testId}:" -BackgroundColor Green -ForegroundColor Black

$testId = "{B052D220-F55E-0DA8-056A-F8753B893755}{1}{B0}"

$section = Get-OneNoteSectionCollectionSection -SectionCollection $sectionCollection -ID $testId
#$section
Write-Host "Get Pages for Section {$testId}:" -BackgroundColor Green -ForegroundColor Black
$pageCollection = Get-OneNoteSectionPageCollection -Section $section
#$pageCollection
######################################################################################

# test

Write-Host "Enriched Page Collection: " -BackgroundColor Green -ForegroundColor Black
$pages = $null
$pages = Get-OneNoteEnrichPageCollection -PageCollection $pageCollection
$pages
#$global:error
#$error.Clear()
<#
$id = $pageCollection[0].ID
$testOutput = "c:\temp\test.docx"

######################################################################################
# Publish a page to .one format (0)
######################################################################################
Write-Host "Publish Page {$testPageId} to .one format:" -BackgroundColor Green -ForegroundColor Black
$testOutput = "c:\temp\test.one"
Invoke-OneNotePublish -ID $id -Path $testOutput -PublishFormat 'pfOneNote' -Overwrite $true

######################################################################################
# Publish a page to .onepkg format (1)
######################################################################################
Write-Host "Publish Page {$testPageId} to .onepkg format:" -BackgroundColor Green -ForegroundColor Black
$testOutput = "c:\temp\test.onepkg"
Invoke-OneNotePublish -ID $id -Path $testOutput -PublishFormat 'pfOneNotePackage' -Overwrite $true

######################################################################################
# Publish a page to .mht format (2)
######################################################################################
Write-Host "Publish Page {$testPageId} to .mht format:" -BackgroundColor Green -ForegroundColor Black
$testOutput = "c:\temp\test.mht"
Invoke-OneNotePublish -ID $id -Path $testOutput -PublishFormat 'pfMHTML' -Overwrite $true

######################################################################################
# Publish a page to .pdf format (3)
######################################################################################
Write-Host "Publish Page {$testPageId} to .pdf format:" -BackgroundColor Green -ForegroundColor Black
$testOutput = "c:\temp\test.pdf"
Invoke-OneNotePublish -ID $id -Path $testOutput -PublishFormat 'pfPDF' -Overwrite $true

######################################################################################
# Publish a page to .xps format (4)
######################################################################################
Write-Host "Publish Page {$testPageId} to .xps format:" -BackgroundColor Green -ForegroundColor Black
$testOutput = "c:\temp\test.xps"
Invoke-OneNotePublish -ID $id -Path $testOutput -PublishFormat 'pfXPS' -Overwrite $true

######################################################################################
# Publish a page to .doc or docx format (5)
######################################################################################
$testOutput = "c:\temp\test.docx"
Write-Host "Publish Page {$testPageId} to .docx format:" -BackgroundColor Green -ForegroundColor Black
Invoke-OneNotePublish -ID $id -Path $testOutput -PublishFormat 'pfWord' -Overwrite $true
Write-Host "Publish Page {$testPageId} to .docx format:" -BackgroundColor Green -ForegroundColor Black
######################################################################################
# (helper method:)
Invoke-OneNotePublishPageToWord -ID $id -Path $testOutput -Overwrite $false
######################################################################################
$testOutput = "c:\temp\test.doc"
Write-Host "Publish Page {$testPageId} to .doc format:" -BackgroundColor Green -ForegroundColor Black
Invoke-OneNotePublish -ID $id -Path $testOutput -PublishFormat 'pfWord' -Overwrite $true

######################################################################################
# Publish a page to .emf format (6)
######################################################################################
Write-Host "Publish Page {$testPageId} to .emf format:" -BackgroundColor Green -ForegroundColor Black
$testOutput = "c:\temp\test.emf"
Invoke-OneNotePublish -ID $id -Path $testOutput -PublishFormat 'pfEMF' -Overwrite $true

######################################################################################
# Publish a page to .HTML format (7) (note will produce .htm not .html)
######################################################################################
Write-Host "Publish Page {$testPageId} to .html format:" -BackgroundColor Green -ForegroundColor Black
$testOutput = "c:\temp\test.htm"
Invoke-OneNotePublish -ID $id -Path $testOutput -PublishFormat 'pfHTML' -Overwrite $true

######################################################################################
# Publish a page to 2007 .one format format (8) (note that this will close the client)
######################################################################################
#
# This resets the client so irritating to test:
#
#Write-Host "Publish Page {$testPageId} to .one format:" -BackgroundColor Green -ForegroundColor Black
#$testOutput = "c:\temp\test2007.one"
#Invoke-OneNotePublish -ID $id -Path $testOutput -PublishFormat 'pfOneNote2007' -Overwrite $true
#>