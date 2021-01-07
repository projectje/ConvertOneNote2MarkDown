<#
    .SYNOPSIS
        Operations that operate on a single OneNote Notebook
        https://docs.microsoft.com/en-us/javascript/api/onenote/onenote.notebook?view=onenote-js-1.1
#>
Import-Module "$PSScriptRoot\FileOperations.psm1" -Force

function Get-OneNoteNotebook {
    <#
        .SYNOPIS
            Returns onenote notebook by id (name, nickname, ID, path, lastModifiedTime, color, isCurrentlyViewed, isUnread, Section, SectionGroup)
    #>
    param(
        [System.Xml.XmlElement]$NotebookCollection,
        [int]$NotebookItem
    )
    return $NotebookCollection.Notebook[$NotebookItem]
}

function Get-OneNoteNotebookSectionCollection {
    <#
        .SYNOPIS
            Returns Sections in a notebook
    #>
    param(
        [System.Xml.XmlElement]$Notebook
    )
    return $Notebook.Section
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
    $sectionGroupCollection = $Notebook.SectionGroup
    if ($Include_Recyclebin -eq $false) {
        $sectionGroupCollection = $sectionGroupCollection | Where-Object { $_.name -ne "OneNote_RecycleBin"}
    }
    return $sectionGroupCollection
}

function Get-OneNoteNotebookCleanFileName {
    <#
        .SYNOPSIS
            returns a name of a notebook to be used as OS directory
    #>
    param(
        [System.Xml.XmlElement]$Notebook
    )
    return $Notebook.Name | Remove-InvalidFileNameChars
}