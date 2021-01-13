<#
    .SYNOPSIS
        Operations on SectionGroup Collections:
            - can contain sections
            - can contain sectiongroups
#>

function Get-OneNoteSectionGroupSectionCollection {
    <#
        .SYNOPIS
            Returns from one section group one section collection
    #>
    param(
        [System.Array]$SectionGroup,
        [string]$ID
    )
    try {
        $section = Get-OneNoteSectionGroupCollectionSectionCollection -SectionGroup $SectionGroup | Where-Object { $_.ID -eq $ID}
        return $section.Section
    }
    catch {
        Throw
    }
}

function Get-OneNoteSectionGroupSectionGroupCollection {
    <#
        .SYNOPSIS
            Returns from one section group one sectiongroup collection
    #>
    param(
        [System.Array]$SectionGroup,
        [string]$ID
    )
    try {
        $sectionGroup = Get-OneNoteSectionGroupCollectionSectionGroupCollection -SectionGroup $SectionGroup | Where-Object { $_.ID -eq $ID}
        return $sectionGroup.Sectiongroup
    }
    catch {
        Throw
    }
}