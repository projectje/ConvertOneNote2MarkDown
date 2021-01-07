<#
    .SYNOPSIS
        operations on a OneNote notebook collections
#>

function Get-OneNoteNotebookCollection {
    <#
        .SYNOPSIS
            Gets the collection of notebooks that are open in the OneNote application instance.
            https://docs.microsoft.com/en-us/javascript/api/onenote/onenote.notebookcollection?view=onenote-js-1.1
    #>
    param(
        [System.Array]$Hierarachy
    )
    return $Hierarchy.Notebooks
}

function Get-OneNoteNotebookCollectionCount {
    param (
        [System.Xml.XmlElement]$NotebookCollection
    )
    return $NotebookCollection.Notebook.Count
}