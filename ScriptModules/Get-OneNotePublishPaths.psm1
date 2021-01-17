<#
    .SYNOPSIS
        Enriches a page object with needed output paths

        ScriptModule: parent is responsible for including necessary modules
#>

function Get-OneNotePublishPaths {
    <#
        Helper for publish: enriched Page object with a certain path set
        https://docs.microsoft.com/en-us/office/client-developer/onenote/enumerations-onenote-developer-reference#odc_PublishFormat
    #>
    param(
        [PSCustomObject]$Config,
        [PSCustomObject]$EnrichedPageObject
    )
    try {
        foreach ($exportFormat in $Config.ExportFormat) {
            $ExportFormat = $exportFormat.Trim() # just in case
            $Extension = $ExportFormat

            if ($Config.PandocMdFormats -contains $ExportFormat) {
                $Extension = 'md'
                # it is highly logical to have a central place for images when exporting to multiple md types at once:
                if ($Config.MdCentralMediaPath -eq $true) {
                    $mediaPath = (Join-Path -Path $Config.ExportRootPath -ChildPath $ExportFormat | Join-Path -ChildPath $Config.Notebook)
                    $EnrichedPageObject | Add-Member -Type NoteProperty -Name 'MediaPath' -Value $mediaPath -Force
                }
                else {
                    $mediaPath = (Join-Path -Path $Config.ExportRootPath -ChildPath $ExportFormat | Join-Path -ChildPath $EnrichedPageObject.RelativePath | Join-Path -ChildPath $EnrichedPageObject.FullName)
                    $EnrichedPageObject | Add-Member -Type NoteProperty -Name 'MediaPath' -Value $mediaPath -Force
                }
            }

            $path = (Join-Path -Path $Config.ExportRootPath -ChildPath $ExportFormat | Join-Path -ChildPath $EnrichedPageObject.RelativePath | Join-Path -ChildPath $EnrichedPageObject.FullName) + "." + $Extension
            $EnrichedPageObject | Add-Member -Type NoteProperty -Name $ExportFormat -Value $path -Force
        }
        return $EnrichedPageObject
    }
    catch {
        throw
    }
}
