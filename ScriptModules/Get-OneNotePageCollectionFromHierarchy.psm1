<#
    .SYNOPSIS
        returns an enriched page collection for all notebooks

        ScriptModule: parent is responsible for including necessary modules
#>
function Get-OneNotePageCollectionFromSectionCollection {
    <#
        Returns page collection from Collection of Sections (each holding a page collection)
    #>
    param(
        [String]$RelativePath,
        [System.Array] $SectionCollection,
        [hashtable]$TagsTable
    )
    try {
        $pageCollection = New-Object -TypeName "System.Collections.ArrayList"
        foreach ($section in $SectionCollection) {
            $sectionFileName = $section.Name | Remove-InvalidFileNameChars
            $NewRelativePath = Join-Path $RelativePath -ChildPath $sectionFileName

            Write-Host "Fetching pages from Section $NewRelativePath" -ForegroundColor Green
            [System.Array]$returnedPageCollection = Get-OneNoteSectionPageCollection -Section $section
            # outputting collections causes powershell to enumerate them by default, always returning powershell array
            [array]$enrichedPageCollection = Get-OneNoteEnrichPageCollection -PageCollection $returnedPageCollection -RelativePath $NewRelativePath
            [System.Collections.ArrayList]$enrichedPageCollectionArrayList = $enrichedPageCollection
            if ($null -ne $enrichedPageCollectionArrayList) {
                # page collection
                $pageCollection.AddRange($enrichedPageCollectionArrayList)

                # tags
                foreach($page in $enrichedPageCollectionArrayList) {
                    $TagsTable = Get-OneNoteTags -EnrichedPageObject $page -TagsTable $TagsTable
                }
            }
        }
        return $pageCollection, $TagsTable
    }
    catch {
        throw
    }
}

function Get-OneNotePageCollectionFromSectionGroupSection {
    <#
        Returns page collection from One item in a SectionGroup Collection which can be a sectiongroup collection or a section collection
    #>
    param (
        [String]$RelativePath,
        [System.Xml.XmlElement]$SectionGroup,
        [string]$Typez,
        [hashtable]$TagsTable
    )
    try {
        $sectionGroupName = $SectionGroup.name | Remove-InvalidFileNameChars
        $NewRelativePath = Join-Path $RelativePath -ChildPath $sectionGroupName
        if ($Typez -eq "Section") {
            return Get-OneNotePageCollectionFromSectionCollection -SectionCollection $SectionGroup.Section -RelativePath $NewRelativePath -TagsTable $TagsTable
        }
        elseif ($Typez -eq "SectionGroup") {
            return Get-OneNotePageCollectionFromSectionGroupCollection -SectionGroupCollection $SectionGroup.SectionGroup -RelativePath $NewRelativePath -TagsTable $TagsTable
        }
    }
    catch {
        throw
    }
}

function Get-OneNotePageCollectionFromSectionGroupCollection {
    <#
        Returns page collection from SectionGroup Collection. A section group can contain an entry "Section" or an entry "SectionGroup"
    #>
    param (
        [String]$RelativePath,
        [System.Array]$SectionGroupCollection,
        [hashtable]$TagsTable
    )
    try {
        if ($SectionGroupCollection.Count -gt 0 -and $null -ne $SectionGroupcollection) {
            $pageCollection = New-Object -TypeName "System.Collections.ArrayList"
            Write-Host "Fetching pages from New SectionGroup $RelativePath" -ForegroundColor Green
            $sectionGroupItemsWithSection = Get-OneNoteSectionGroupCollectionSectionCollection -SectionGroup $SectionGroupCollection
            Write-Host "Fetching pages from Entries Section in SectionGroup" -ForegroundColor Green
            foreach ($sectionGroup in $sectionGroupItemsWithSection) {
                $returnedPageCollection, $TagsTable = Get-OneNotePageCollectionFromSectionGroupSection -SectionGroup $sectionGroup -Typez "Section" -RelativePath $RelativePath -TagsTable $TagsTable
                $pageCollection.AddRange($returnedPageCollection)
            }
            # get all entries that contain an entry "sectiongroup"
            Write-Host "Fetching pages from Entries SectionGroup in SectionGroup" -ForegroundColor Green
            $sectionGroupItemsWithSectionGroup = Get-OneNoteSectionGroupCollectionSectionGroupCollection -SectionGroup $sectionGroupCollection
            foreach ($sectionGroup in $sectionGroupItemsWithSectionGroup ) {
                $returnedPageCollection, $TagsTable = Get-OneNotePageCollectionFromSectionGroupSection -SectionGroup $sectionGroup -Typez "SectionGroup" -RelativePath $RelativePath -TagsTable $TagsTable
                $pageCollection.AddRange($returnedPageCollection)
            }
            return $pageCollection, $TagsTable
        }

    }
    catch {
        throw
    }
}

function Get-OneNotePageCollectionFromNotebook {
    <#
        Returns page collection from one OneNote Notebook (which contains sections and sectiongroups)
    #>
    param (
        [System.Xml.XmlElement]$Notebook,
        [hashtable]$TagsTable
    )
    try {
        $pageCollection = New-Object -TypeName "System.Collections.ArrayList"
        Write-Host "Fetching pages from Sections and SectionGroups from Notebook $notebookFileName" -ForegroundColor Green
        $notebookFileName = $Notebook.Name | Remove-InvalidFileNameChars
        $sectionCollection = Get-OneNoteNotebookSectionCollection -Notebook $Notebook
        $sectionGroupCollection = Get-OneNoteNotebookSectionGroupCollection -Notebook $Notebook -Include_Recyclebin $false
        if ($null -ne $sectionCollection) {
            $returnedPageCollection, $TagsTable = Get-OneNotePageCollectionFromSectionCollection -SectionCollection $sectionCollection -RelativePath $notebookFileName -TagsTable $TagsTable
            $pageCollection.AddRange($returnedPageCollection)
        }
        if ($null -ne $sectionGroupCollection) {
            $returnedPageCollection, $TagsTable = Get-OneNotePageCollectionFromSectionGroupCollection -SectionGroupCollection $sectionGroupCollection -RelativePath $notebookFileName -TagsTable $TagsTable
            $pageCollection.AddRange($returnedPageCollection)
        }
        return $pageCollection, $TagsTable
    }
    catch {
        throw
    }
}

function Get-OneNotePageCollectionFromNotebookCollection {
    <#
       Returns page collection from all Notebooks
    #>
    param (
        [System.Xml.XmlElement]$NotebookCollection,
        [hashtable]$TagsTable
    )
    try {
        $pageCollection = New-Object -TypeName "System.Collections.ArrayList"
        $count = Get-OneNoteNotebookCollectionCount -NotebookCollection $NotebookCollection
        if ($count -eq 0) {
            Write-Host "Warning: no notebooks found"
        }
        for ($item = 0; $item -lt $count; $item++) {
            Write-Host "Fetching pages from Notebook $($item+1) out of $count" -ForegroundColor Green
            $notebook = Get-OneNoteNotebook -NotebookCollection $NotebookCollection -NotebookItem $item
            $returnedPageCollection, $TagsTable = Get-OneNotePageCollectionFromNotebook -Notebook $notebook -TagsTable $TagsTable
            $pageCollection.AddRange($returnedPageCollection)
        }
        return $pageCollection, $TagsTable
    }
    catch {
        throw
    }
}

function Get-OneNotePageCollectionFromHierarchy {
    <#
        Returns page collection from OneNote Hierarchy of all notebooks
    #>
    try {
        $hierarchy = Get-OneNoteHierarchy
        $notebookCollection = Get-OneNoteNoteBookCollection -Hierarchy $hierarchy
        $TagsTable = @{}
        return Get-OneNotePageCollectionFromNotebookCollection -NotebookCollection $notebookCollection -TagsTable $TagsTable
    }
    catch {
       throw
    }
}
