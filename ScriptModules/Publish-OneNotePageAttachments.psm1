<#
    .SYNOPSIS
        publishes all attachments for a certain enriched page object

        ScriptModule: parent is responsible for including necessary modules
#>
function Publish-OneNotePageAttachments {
    <#
        .SYNOPSIS
            Exports all attachments of a page
    #>
    param(
        [PSCustomObject]$Config,
        [PSCustomObject]$EnrichedPageObject
    )
    try {
        $attachmentsPath = Join-Path -Path ($Config.ExportRootPath) -ChildPath "_attachments" | Join-Path -ChildPath $EnrichedPageObject.RelativePath | Join-Path -ChildPath $EnrichedPageObject.FullName
        return Get-OneNotePageInsertedFileObjects -Id $EnrichedPageObject.Id -AttachmentsPath $attachmentsPath -OverwriteAttachments $Config.OverWriteAttachments
    }
    catch {
        throw
    }
}
