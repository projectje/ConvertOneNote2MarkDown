<#
    Operations using OneNote.Application.GetContents
#>
function Get-OneNotePageInsertedFileObjects {
    <#
        export inserted file objects, removing any escaped symbols from filename so that links to them actually work
    #>
    param(
        [string]$PageId,
        [System.Xml.XmlElement]$Page
    )
    try {
        $OneNote = New-Object -ComObject OneNote.Application
        [xml]$pagexml = $null
        $OneNote.GetPageContent($PageId, [ref]$pagexml, 7)
        $fileObjects = $pagexml.Page.Outline.OEChildren.OE | Where-Object { $Page.InsertedFile }
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($OneNote)
        Remove-Variable OneNote
        return $fileObjects
    }
    catch {
        Write-Host $global:error -ForegroundColor Red
        Exit
    }
}