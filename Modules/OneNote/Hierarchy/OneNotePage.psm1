<#
    .SYNOPSIS
        Operations on OneNote Pages
        See: https://docs.microsoft.com/en-us/office/client-developer/onenote/application-interface-onenote
#>


function Get-OneNotePageName {
    <#
        .SYNOPSIS
            Returns the sectiongroup name stripped off invalid chars
    #>
    param(
        [System.Xml.XmlElement]$Page
    )
    try {
        return $Page.name | Remove-InvalidFileNameChars
    }
    catch {
        Write-Host $global:error -ForegroundColor Red
        Exit
    }
}