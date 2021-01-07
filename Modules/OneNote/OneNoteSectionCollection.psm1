function Get-OneNoteSectionCollectionSection {
    <#
        .SYNOPSIS
            Returns on section from a sectioncollection
    #>
    param(
        [System.Array]$SectionCollection,
        [string]$ID
    )
    return $SectionCollection | Where-Object { $_.ID -eq $ID }
}
