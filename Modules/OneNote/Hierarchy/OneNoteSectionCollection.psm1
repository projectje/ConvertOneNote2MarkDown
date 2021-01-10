function Get-OneNoteSectionCollectionSection {
    <#
        .SYNOPSIS
            Returns on section from a sectioncollection
    #>
    param(
        [System.Array]$SectionCollection,
        [string]$ID
    )
    try {
        return $SectionCollection | Where-Object { $_.ID -eq $ID }
    }
    catch {
        Write-Host $global:error -ForegroundColor Red
        Exit
    }
}
