<#
    .SYNOPSIS
        Operations that operate on a single OneNote Notebook
        https://docs.microsoft.com/en-us/javascript/api/onenote/onenote.notebook?view=onenote-js-1.1
#>

function Get-OneNoteNotebook {
    <#
        .SYNOPIS
            Returns onenote notebook by id (name, nickname, ID, path, lastModifiedTime, color, isCurrentlyViewed, isUnread, Section, SectionGroup)
    #>
    param(
        [System.Xml.XmlElement]$NotebookCollection,
        [int]$NotebookItem
    )
    try {
        return $NotebookCollection.Notebook[$NotebookItem]
    }
    catch {
        Write-Host $global:error -ForegroundColor Red
        Exit
    }
}

function Get-OneNoteNotebookSectionCollection {
    <#
        .SYNOPIS
            Returns Sections in a notebook
    #>
    param(
        [System.Xml.XmlElement]$Notebook
    )
    try{
        return $Notebook.Section
    }
    catch {
        Write-Host $global:error -ForegroundColor Red
        Exit
    }
}

function Get-OneNoteNotebookSectionGroupCollection {
    <#
        .SYNOPSIS
            Returns SectionGroups in a notebook
    #>
    param(
        [System.Xml.XmlElement]$Notebook,
        [bool]$Include_Recyclebin = $false
    )
    try {
        $sectionGroupCollection = $Notebook.SectionGroup
        if ($Include_Recyclebin -eq $false) {
            $sectionGroupCollection = $sectionGroupCollection | Where-Object { $_.name -ne "OneNote_RecycleBin"}
        }
        return $sectionGroupCollection
    }
    catch
    {
        Write-Host $global:error -ForegroundColor Red
        Exit
    }
}
