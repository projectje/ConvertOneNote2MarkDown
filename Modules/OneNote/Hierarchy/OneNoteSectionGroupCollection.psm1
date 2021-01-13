<#
    .SYNOPSIS
        Operations on SectionGroup Collections:
            - can contain sections
            - can contain sectiongroups
#>
function Get-OneNoteSectionGroupCollectionSectionGroupCollection {
    <#
        .SYNOPSIS
            Returns all section group collection in a section group collection
    #>
    param(
        [System.Array]$SectionGroupCollection
    )
    try {
        return $SectionGroupCollection | Where-Object {$null -ne $_.SectionGroup}
    }
    catch {
        Throw
    }
}

function Get-OneNoteSectionGroupCollectionSectionCollection {
    <#
        .SYNOPSIS
            Returns all sectioncollection in a section group collection
    #>
    param(
        [System.Array]$SectionGroupCollection
    )
    try {
        return $SectionGroupCollection |Where-Object {$null -ne $_.Section}
    }
    catch
    {
        Throw
    }
}
