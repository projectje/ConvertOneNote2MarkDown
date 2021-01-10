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
        Write-Host $global:error -ForegroundColor Red
        Exit
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
        Write-Host $global:error -ForegroundColor Red
        Exit
    }
}
