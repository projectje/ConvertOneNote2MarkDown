<#
    .SYNOPSIS
        publishes complete onenotehierarchy in pages, attachments, etc via configuration in configuration file

        ScriptModule: parent is responsible for including necessary modules
#>
function Publish-OneNoteHierarchyPages {
    param (
        [PSCustomObject]$Config
    )
    $Config

    $enrichedPageCollection, $tagsTable = Get-OneNotePageCollectionFromHierarchy
    Write-Host "Done fetching pages. Fetched $($pageCollection.count) pages" -ForegroundColor Green
    Publish-OneNoteTags -Config $Config -TagsTable $tagsTable
    foreach ($page in $enrichedPageCollection) {
        $OutputEncoding = [ System.Text.Encoding]::UTF8
        $attachments = Publish-OneNotePageAttachments -Config $Config -EnrichedPageObject $page
        $EnrichedPageObject = Get-OneNotePublishPaths -Config $Config -EnrichedPageObject $page
        $Config.ExportFormat | ForEach-Object {
            Invoke-OneNotePublish -Id $EnrichedPageObject.Id -PublishFormat $_ -Path ($EnrichedPageObject | Select-Object -ExpandProperty $_) -Overwrite $Config.Overwrite
        }
        ConvertTo-MdFromDocx -Config $Config -EnrichedPageObject $EnrichedPageObject
    }
}