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
    try {
        return $Section.Page
    }
    catch {
        Throw
    }
}

function Get-OneNoteSectionPageCount {
    <#
        Returns page count in a section
    #>
    param(
        [System.Xml.XmlElement]$Section
    )
    try {
        return $Section.Page.Count
    }
    catch {
        Throw
    }
}
