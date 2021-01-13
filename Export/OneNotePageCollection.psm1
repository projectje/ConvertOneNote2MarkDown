<#
    Child Module (parents needs to reference the dependent modules)
#>

function Publish-PageAttachments {
    <#
        .SYNOPSIS
            Exports all attachments of a page
    #>
    param(
        [object]$Config,
        [String]$RelativePath,
        [object]$Page
    )
    try {
        $attachmentsPath = Join-Path -Path ($Config.ExportRootPath) -ChildPath "attachments" | Join-Path -ChildPath $RelativePath | Join-Path -ChildPath $Page.FullName
        return Get-OneNotePageInsertedFileObjects -Id $Page.ID -AttachmentsPath $attachmentsPath -OverwriteAttachments $Config.OverWriteAttachments
    }
    catch {
        throw
    }
}

function Get-EnrichedPage {
    <#
        .SYNOPSIS
            Enriched the page object with document paths
    #>
    param (
        [object]$Config,
        [String]$RelativePath,
        [object]$Page
    )
    try {
        foreach($exportFormat in $Config.ExportFormat) {
            $Page = Get-EnrichedPagePublishPaths -Page $Page -ExportFormat $exportFormat -Config $Config -RelativePath $RelativePath
        }
        return $Page
    }
    catch {
        throw
    }
}

function Publish-OneNotePage {
    <#
        .SYNOPIS
            handle onenote publishing
    #>
    param (
        [object]$Config,
        [object]$Page
    )
    try {
        foreach($exportFormat in $Config.ExportFormat) {
            if ($Config.PublishFormats -contains $exportFormat) {
                $documentOutputPath = $Page | Select-Object -ExpandProperty $exportFormat
                Invoke-OneNotePublish -Id $Page.Id -Path $documentOutputPath -PublishFormat $exportFormat -Overwrite $Config.Overwrite
            }
        }
    }
    catch {
        throw
    }
}

function Get-EnrichedPagePublishPaths {
    <#
        Helper for publish: enriched Page object with a certain path set
        https://docs.microsoft.com/en-us/office/client-developer/onenote/enumerations-onenote-developer-reference#odc_PublishFormat
    #>
    param(
        [object]$Config,
        [string]$RelativePath,
        [object]$Page,
        [string]$ExportFormat
    )
    try {
        $ExportFormat = $ExportFormat.Trim()
        $Extension = $ExportFormat

        if ($Config.PandocMdFormats -contains $ExportFormat) {
            $Extension = 'md'
            # md files are converted from docx files. These contain, when unzipped a folder "media"
            # optional is to make one central media location per notebook:
            if ($Config.MdCentralMediaPath -eq $true)
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

        $path = (Join-Path -Path ($Config.ExportRootPath) -ChildPath $ExportFormat | Join-Path -ChildPath $RelativePath | Join-Path -ChildPath ($Page.FullName)) + "." + $Extension
        $Page | Add-Member -Type NoteProperty -Name $ExportFormat -Value $path -Force
        return $Page
    }
    catch {
        throw
    }
}


function Invoke-PandocConversion {
    <#
        .SYNOPSIS
            Handle Pandoc Conversion
    #>
    param (
        [object]$Config,
        [String]$RelativePath,
        [object]$Page
    )
    try {
        # once docx has been created handle pandoc MD conversions
        # todo: if specifying multiple md formats media will be exported multiple times
        foreach($exportFormat in $Config.ExportFormat) {
            if ($Config.PandocMdFormats -contains $exportFormat) {
                $documentInputPath = $Page | Select-Object -ExpandProperty 'docx'
                $documentOutputPath = $Page | Select-Object -ExpandProperty $exportFormat
                New-Dir -Path  ([IO.Path]::GetDirectoryName($documentOutputPath)) | Out-Null
                Invoke-ConvertDocxToMd -PandocExec $Config.Pandoc -OutputFormat $exportFormat -Inputfile $documentInputPath -OutputFile $documentOutputPath -MediaPath $Page.MediaPath

                # filters:
                if ($Config.MdClearSpaces -eq $true) {
                    #    Invoke-MdClearSpaces -MdPath $documentOutputPath
                }
                if ($Config.MdClearEspace -eq $true) {
                    #    Invoke-MdClearEscape -MdPath $documentOutputPath
                }
                if ($Config.MdAddYaml -eq $true) {
                    #    Invoke-MdAddYaml -MdPath $documentOutputPath -PageName $Page.Name -PageDateTime $Page.DateTime
                }
                Invoke-MdRenameImages -MdPath $documentOutputPath -MediaPath $Page.MediaPath -PageName $Page.Name
                #Invoke-MdImagePathReference -MdPath $documentOutputPath -MediaPath $Page.MediaPath -LevelsPrefix $Level
            }
        }
    }
    catch {
        throw
    }
}

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
    try
    {
        [array]$pages = Get-OneNoteEnrichPageCollection -PageCollection $PageCollection
        foreach ($page in $pages) {
            $attachments = Publish-PageAttachments -Config $Config -RelativePath $RelativePath -Page $page
            $publishPage = Get-EnrichedPage -Config $Config -RelativePath $RelativePath -Page $Page
            Publish-OneNotePage -Config $Config -Page $publishPage
            Invoke-PandocConversion -Config $Config -RelativePath $RelativePath -Page $publishPage
        }
    }
    catch {
        throw
    }
}