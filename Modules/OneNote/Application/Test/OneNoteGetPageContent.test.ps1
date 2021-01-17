Import-Module "$PSScriptRoot\..\..\Hierarchy\OneNoteHierarchy.psm1" -Force
Import-Module "$PSScriptRoot\..\..\Hierarchy\OneNoteNotebookCollection.psm1" -Force
Import-Module "$PSScriptRoot\..\..\Hierarchy\OneNoteNotebook.psm1" -Force
Import-Module "$PSScriptRoot\..\..\Hierarchy\OneNoteSectionCollection.psm1" -Force
Import-Module "$PSScriptRoot\..\..\Hierarchy\OneNoteSection.psm1" -Force
Import-Module "$PSScriptRoot\..\..\Hierarchy\OneNotePageCollection.psm1" -Force
Import-Module "$PSScriptRoot\..\..\Hierarchy\OneNotePage.psm1" -Force
Import-Module "$PSScriptRoot\..\OneNoteGetPageContent.psm1" -Force


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

#$hash = @{Key = "to do"; Value= "ss"}
#if ($hash.Key -contains"to d") {
#    write-host " yesy "
#}


$xml = Get-OneNotePageXML -Id $pageCollection[15].ID

if ($null -ne $xml.TagDef) {
    $tagArray = @()
    $countTags = $xml.TagDef.count
    for ($i=0;$i -lt $countTags; $i++) {
        $tag = [PSCustomObject]@{
            index = $xml.TagDef[$i].index
            type = $xml.TagDef[$i].type
            symbol = $xml.TagDef[$i].symbol
            fontColor = $xml.TagDef[$i].fontcolor
            highlightColor = $xml.TagDef[$i].highlightcolor
            name = $xml.TagDef[$i].name
        }
        $tagArray += $tag
    }
}

$countchildren = $xml.Outline.OEChildren.OE.count
for ($i=0; $i -lt $countchildren; $i++) {
    # get item, if item is tag then retrieve the next item
    # get item by object id GetBinaryPageContent
    if ($null -ne $xml.Outline.OEChildren.OE[$i].Tag) {

        $xml.Outline.OEChildren.OE[$i].Tag
        $data = $xml.Outline.OEChildren.OE[$i].T."#cdata-section"


        # determine the pagename to write to
        $taskname = $tagArray | Where-Object -Property index -eq $xml.Outline.OEChildren.OE[$i].Tag.index | Select-Object name
        if ($xml.Outline.OEChildren.OE[$i].Tag.completed -eq 'true') {
            write-host $taskname.name + "_completed" + $data
        } else {
            write-host $taskname.name  + $data
        }

        # determine the line to add
       # $xml.Outline.OEChildren.OE[$i].Tag

        # $xml.Outline.OEChildren.OE[$i].Tag.index
        # $xml.Outline.OEChildren.OE[$i].Tag.completed
        # $xml.Outline.OEChildren.OE[$i].Tag.disabled
        # $xml.Outline.OEChildren.OE[$i].Tag.creationDate
        # $xml.Outline.OEChildren.OE[$i].Tag.completionDate



        # $tagArray | Where-Object -Property index -eq 1
    }
}

#$pageId = $pageCollection[15].ID
#$objectId = $xml.Outline.OEChildren.OE[10].objectID

#$pageId
#$xml.Outline.OEChildren.OE[10]

#$binarycontent = Get-OneNoteBinaryPageContent -PageId pageId -ObjectId $objectId

# first check if $xml.TagDef exists and if the name is "To Do"

#if ($null -ne $xml.TagDef) {
#    $count = $xml.TagDef.Count
    # create a subdir per tag and per tag if it is complete or not
    # //notebook/_tags/tagname_completed.txt
    # //notebook/_tags/tagname_notcompleted.txt
    # if the file does not exist create it
    # if the file exists append to it ... however... we need the line following this AND the complete topic page
#}
