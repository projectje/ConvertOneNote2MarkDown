<#
    Child Module (parents needs to reference the dependent modules)
#>
function Export-OneNotePageCollection {
    <#
        Exports all pages in a page collection
    #>
    param(
        [Object]$Config,
        [String]$RelativePath,
        [System.Array]$PageCollection,
        [Int]$Level
    )
    try {
        [array]$pages = Get-OneNoteEnrichPageCollection -PageCollection $PageCollection
        [array]$publishformats = Get-OneNotePublishFormats
        [array]$pandocMdFormats = Get-PandocMDOutputFormats

        foreach ($page in $pages) {
            # get output paths
            $publishPage = $page
            $Config.ExportFormat -split ',' -replace '^\s+|\s+$' | ForEach-Object {
                $publishPage = Get-EnrichedPagePublishPaths -Page $publishPage -ExportFormat $_ -Config $Config -RelativePath $RelativePath
            }

            # handle onenote publishing
            $Config.ExportFormat -split ',' -replace '^\s+|\s+$' | ForEach-Object {
                if ($publishformats -contains $_) {
                    $documentOutputPath = $publishPage | Select-Object -ExpandProperty $_
                    Invoke-OneNotePublish -ID $publishPage.Id -Path $documentOutputPath -PublishFormat $_ -Overwrite ([System.Convert]::ToBoolean($Config.Overwrite))
                    $attachmentsPath = $publishPage | Select-Object -ExpandProperty "$($_)AttachmentsPath"
                    $result = Get-OneNotePageInsertedFileObjects -ID $publishPage.Id -AttachmentsPath $attachmentsPath
                }
            }

            # handle pandoc MD conversions
            $Config.ExportFormat -split ',' -replace '^\s+|\s+$' | ForEach-Object {
                if ($pandocMdFormats -contains $_) {
                    $documentInputPath = $publishPage | Select-Object -ExpandProperty 'docx'
                    $documentOutputPath = $publishPage | Select-Object -ExpandProperty $_
                    New-Dir -Path  ([IO.Path]::GetDirectoryName($documentOutputPath)) | Out-Null
                    Invoke-ConvertDocxToMd -PandocExec $Config.Pandoc -OutputFormat $_ -Inputfile $documentInputPath -OutputFile $documentOutputPath -MediaPath $publishPage.MediaPath
                    # filters:
                    if ([bool]($Config.PSobject.Properties.name -match "MdClearSpaces")) {
                        Invoke-MdClearSpaces -MdPath $documentOutputPath
                    }
                    if ([bool]($Config.PSobject.Properties.name -match "MdClearEspace")) {
                        Invoke-MdClearEscape -MdPath $documentOutputPath
                    }
                    if ([bool]($Config.PSobject.Properties.name -match "MdAddYaml")) {
                        Invoke-MdAddYaml -MdPath $documentOutputPath -PageName $publishPage.Name -PageDateTime $page.DateTime
                    }
                    Invoke-MdRenameImages -MdPath $documentOutputPath -MediaPath $publishPage.MediaPath -PageName $publishPage.Name
                    Invoke-MdImagePathReference -MdPath $documentOutputPath -MediaPath $publishPage.MediaPath -LevelsPrefix $Level
                }
            }
        }


    }
    catch {
        Write-Host $global:error -ForegroundColor Red
        $global:Error
        Exit
    }
}

function Get-EnrichedPagePublishPaths {
    <#
        Helper for publish: enriched Page object with a certain path set
        https://docs.microsoft.com/en-us/office/client-developer/onenote/enumerations-onenote-developer-reference#odc_PublishFormat
    #>
    param(
        [Object]$Config,
        [string]$RelativePath,
        [object]$Page,
        [string]$ExportFormat
    )
    try {
        $ExportFormat = $ExportFormat.Trim()
        $Extension = $ExportFormat

        [array]$pandocMdFormats = Get-PandocMDOutputFormats
        if ($pandocMdFormats -contains $ExportFormat) {
            $Extension = 'md'
            # md files are converted from docx files. These contain, when unzipped a folder "media"
            # optional is to make one central media location per notebook:
            if ([bool]($Config.PSobject.Properties.name -match "MdCentralMediaPath"))
            {
                $mediaPath = (Join-Path -Path ($Config.ExportRootPath) -ChildPath $ExportFormat | Join-Path -ChildPath ($Config.Notebook))
                $Page | Add-Member -Type NoteProperty -Name 'MediaPath' -Value $mediaPath -Force
            }
            else
            {
                $mediaPath = (Join-Path -Path ($Config.ExportRootPath) -ChildPath $ExportFormat | Join-Path -ChildPath $RelativePath | Join-Path -ChildPath $Page.FullName)
                $Page | Add-Member -Type NoteProperty -Name 'MediaPath' -Value $mediaPath -Force
            }
        }

        $attachmentsPath = (Join-Path -Path ($Config.ExportRootPath) -ChildPath $ExportFormat | Join-Path -ChildPath $RelativePath | Join-Path -ChildPath $Page.FullName) | Join-Path -ChildPath "Attachments"
        $Page | Add-Member -Type NoteProperty -Name "$($ExportFormat)AttachmentsPath" -Value $attachmentsPath -Force

        $path = (Join-Path -Path ($Config.ExportRootPath) -ChildPath $ExportFormat | Join-Path -ChildPath $RelativePath | Join-Path -ChildPath ($Page.FullName)) + "." + $Extension
        $Page | Add-Member -Type NoteProperty -Name $ExportFormat -Value $path -Force
        return $Page
    }
    catch {
        Write-Host $global:error -ForegroundColor Red
        $global:Error
        Exit
    }
}