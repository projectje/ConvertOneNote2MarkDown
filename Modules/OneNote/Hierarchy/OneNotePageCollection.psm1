
<#
    .SYNOPSIS
        Handles operations on Page Collection
#>
function Get-OneNotePageCollectionCount {
    <#
        Gets the count of pages in a page collection
    #>
    param(
        [System.Array]$PageCollection
    )
    try {
        return $PageCollection.Count
    }
    catch {
        Write-Host $global:error -ForegroundColor Red
        Exit
    }
}

function Get-OneNotePageFromPageCollection {
    <#
        Gets a page by ID from the page collection
    #>
    param(
        [System.Array]$PageCollection,
        [string]$ID
    )
    try {
        return $PageCollection | Where-Object { $_.ID -eq $ID}
    }
    catch {
        Write-Host $global:error -ForegroundColor Red
        Exit
    }
}
