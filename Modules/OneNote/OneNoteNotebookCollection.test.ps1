Import-Module "$PSScriptRoot\OneNoteHierarchy.psm1" -Force
Import-Module "$PSScriptRoot\OneNoteNotebookCollection.psm1" -Force

######################################################################################
# Setup (tested in other tests)
######################################################################################
$hierarchy = Get-OneNoteHierarchy

######################################################################################
# Get-OneNoteNoteBookCollection
# Verify the type
######################################################################################
$notebookCollection = Get-OneNoteNotebookCollection  -Hierarchy $hierarchy
Write-Host "Notebooks should be XmlElement" -BackgroundColor Blue -ForegroundColor White
$notebookstype = $notebookCollection.GetType()
if ($notebookstype.BaseType -eq [System.Xml.XmlLinkedNode])
{
    Write-Host "OK"
}
# Have the property notebook
Write-Host "Notebooks contain property Notebook" -BackgroundColor Blue -ForegroundColor White
foreach ($property in $notebookCollection | get-member) {
    if ($property.MemberType -eq "Property" -and $property.Name -eq "Notebook") {
        Write-Host "OK"
    }
}
# Dump content
Write-Host "Content of notebooks:" -BackgroundColor Green -ForegroundColor Black
$notebookCollection.Notebook

######################################################################################
# Get-OneNoteNoteBookCollectionCount
# Dump count of notebooks
######################################################################################
Write-Host "Count of notebooks:" -BackgroundColor Green -ForegroundColor Black
Get-OneNoteNotebookCollectionCount -NotebookCollection $notebookCollection
