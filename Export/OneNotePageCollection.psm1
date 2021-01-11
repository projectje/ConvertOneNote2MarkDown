<#
    Child Module (parents needs to reference the dependent modules)
#>
function Export-OneNotePageCollection {
    <#
        Exports all pages in a page collection
    #>
    param(
        [Object]$Config,
        [System.Array]$PageCollection,
        [String]$Path,
        [Int]$Level
    )
    try {
        [array]$pages = Get-OneNoteEnrichPageCollection -PageCollection $PageCollection
        [array]$publishformats = Get-OneNotePublishFormats

        # handle onenote publishing
        foreach ($page in $pages) {
            $publishPage = Get-OneNoteEnrichedPage -Path $Path -Page $page -ExportRootPath ($Config.ExportRootPath) -ExportFormat ($Config.ExportFormat)
            $Config.ExportFormat -split ',' -replace '^\s+|\s+$' | ForEach-Object {
                $exportFormat = $_
                if ($publishformats -contains $exportFormat) {
                    $documentPath = $publishPage | Select-Object -ExpandProperty $exportFormat
                    Invoke-OneNotePublish -ID ($publishPage.Id) -Path $documentPath -PublishFormat $exportFormat -Overwrite ([System.Convert]::ToBoolean($Config.Overwrite))
                }
            }
        }

        # handle pandoc conversions
        # New-Dir -Path  ([IO.Path]::GetDirectoryName($path)) | Out-Null # not needed
        # Invoke-ConvertDocxToMd -OutputFormat $converter -Inputfile $wordDocumentPath -OutputFile "$fullfilepathwithoutextension.md" -MediaPath $mediaPath

    }
    catch {
        Write-Host $global:error -ForegroundColor Red
        $global:Error
        Exit
    }
}
