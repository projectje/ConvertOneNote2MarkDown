Import-Module "$PSScriptRoot\FileOperations.psm1" -Force
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
    $section = Get-OneNoteSectionGroupCollectionSectionCollection -SectionGroup $SectionGroup | Where-Object { $_.ID -eq $ID}
    return $section.Section
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
    $sectionGroup = Get-OneNoteSectionGroupCollectionSectionGroupCollection -SectionGroup $SectionGroup | Where-Object { $_.ID -eq $ID}
    return $sectionGroup.Sectiongroup
}

function Get-OneNoteSectionGroupName {
    <#
        .SYNOPSIS
            Returns the sectiongroup name stripped off invalid chars
    #>
    param(
        [System.Xml.XmlElement]$SectionGroup
    )
    return $SectionGroup.Name | Remove-InvalidFileNameChars
}