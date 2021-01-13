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
        [System.Array]$Hierarchy
    )
    try{
        return $Hierarchy.Notebooks
    }
    catch {
        Throw
    }
}

function Get-OneNoteNotebookCollectionCount {
    param (
        [System.Xml.XmlElement]$NotebookCollection
    )
    try {
        return $NotebookCollection.Notebook.Count
    }
    catch {
        Throw
    }
}