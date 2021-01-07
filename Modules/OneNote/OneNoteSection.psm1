Import-Module "$PSScriptRoot\FileOperations.psm1" -Force
<#
<#
    .SYNOPSIS
        Operations on Individual Section
#>

function Get-OneNoteSectionPageCollection {
    <#
        Returns pages in a section
    #>
    param(
        [System.Xml.XmlElement]$Section
    )
    return $Section.Page
}

function Get-OneNoteSectionPageCount {
    <#
        Returns page count in a section
    #>
    param(
        [System.Xml.XmlElement]$Section
    )
    return $Section.Page.Count
}

function Get-OneNoteSectionName {
    <#
        .SYNOPSIS
            Returns the section name stripped off invalid chars
    #>
    param(
        [System.Xml.XmlElement]$Section
    )
    return $Section.Name | Remove-InvalidFileNameChars
}