Import-Module "$PSScriptRoot\FileOperations.psm1" -Force
<#
    .SYNOPIS
        Operations on OneNote Pages
        See: https://docs.microsoft.com/en-us/office/client-developer/onenote/application-interface-onenote
#>
function Invoke-OneNotePagePublish {
    <#
        .SYNOPSIS
            Publishes a OneNote page to a Word Document
            See: https://docs.microsoft.com/en-us/office/client-developer/onenote/enumerations-onenote-developer-reference#odc_PublishFormat
            for publish options (mht, pdf, xps, doc, emf, html)
    #>
    param (
        [string] $ID,
        [string] $Path,
        [bool] $Overwrite = $true
    )
    try {
        $OneNotePage = New-Object -ComObject OneNote.Application

        $OneNotePage | get-member

        if ($Overwrite -eq $true) {
            if ([System.IO.File]::Exists($Path)) {
                Remove-Item -path $Path -Force -ErrorAction SilentlyContinue
            }
            Write-Host "2" $ID $Path
            $OneNotePage.Publish($ID, $Path, "pfWord", "")
            Write-Host "3"
        }
        else {
            if (![System.IO.File]::Exists($Path)) {
                $OneNotePage.Publish($ID, $Path, "pfWord", "")
            }
        }
        Write-Host "4"
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($OneNotePage)
        Write-Host "5"
        Remove-Variable OneNotePage
        Write-Host "6"
    }
    catch {
        Write-Host $Error -ForegroundColor Red
        Exit
    }
}

function Get-OneNotePageInsertedFileObjects {
    <#
        export inserted file objects, removing any escaped symbols from filename so that links to them actually work
    #>
    param(
        [string]$PageId,
        [System.Xml.XmlElement]$Page
    )
    $OneNote = New-Object -ComObject OneNote.Application
    [xml]$pagexml = $null
    $OneNote.GetPageContent($PageId, [ref]$pagexml, 7)
    $fileObjects = $pagexml.Page.Outline.OEChildren.OE | Where-Object { $Page.InsertedFile }
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($OneNote)
    Remove-Variable OneNote
    return $fileObjects
}

function Get-OneNotePageName {
    <#
        .SYNOPSIS
            Returns the sectiongroup name stripped off invalid chars
    #>
    param(
        [System.Xml.XmlElement]$Page
    )
    return $Page.name | Remove-InvalidFileNameChars
}