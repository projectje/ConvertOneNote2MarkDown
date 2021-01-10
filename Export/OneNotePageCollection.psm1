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
        foreach ($page in $pages) {
            $paths = Get-OneNotePagePublishPaths -Config $Config -Path $Path -Page $page

            if ($null -ne $paths.docx) {
                Write-Host "Exporting Page: " $paths.docx -BackgroundColor Green
                Invoke-OneNotePublish -ID ($page.Id) -Path $paths.docx -PublishFormat 'pfWord' -Overwrite ([System.Convert]::ToBoolean($Config.Overwrite))
            }
            if ($null -ne $paths.doc) {
                Write-Host "Exporting Page: " $page.doc -BackgroundColor Green
                Invoke-OneNotePublish -ID ($page.Id) -Path $paths.one -PublishFormat 'pfWord' -Overwrite ([System.Convert]::ToBoolean($Config.Overwrite))
            }
            if ($null -ne $paths.one) {
                Write-Host "Exporting Page: " $paths.one -BackgroundColor Green
                Invoke-OneNotePublish -ID ($page.Id) -Path $paths.one -PublishFormat 'pfOneNote' -Overwrite ([System.Convert]::ToBoolean($Config.Overwrite))
            }
            if ($null -ne $paths.onepkg) {
                Write-Host "Exporting Page: " $paths.onepkg -BackgroundColor Green
                Invoke-OneNotePublish -ID ($page.Id) -Path $paths.onepkg -PublishFormat 'pfOneNotePackage' -Overwrite ([System.Convert]::ToBoolean($Config.Overwrite))
            }
            if ($null -ne $paths.mht) {
                Write-Host "Exporting Page: " $paths.mht -BackgroundColor Green
                Invoke-OneNotePublish -ID ($page.Id) -Path $paths.mht -PublishFormat 'pfMHTML' -Overwrite ([System.Convert]::ToBoolean($Config.Overwrite))
            }
            if ($null -ne $paths.pdf) {
                Write-Host "Exporting Page: " $paths.pdf -BackgroundColor Green
                Invoke-OneNotePublish -ID ($page.Id) -Path $paths.pdf -PublishFormat 'pfPDF' -Overwrite ([System.Convert]::ToBoolean($Config.Overwrite))
            }
            if ($null -ne $paths.xps) {
                Write-Host "Exporting Page: " $paths.xps -BackgroundColor Green
                Invoke-OneNotePublish -ID ($page.Id) -Path $paths.xps -PublishFormat 'pfpXPS' -Overwrite ([System.Convert]::ToBoolean($Config.Overwrite))
            }
            if ($null -ne $paths.emf) {
                Write-Host "Exporting Page: " $paths.emf -BackgroundColor Green
                Invoke-OneNotePublish -ID ($page.Id) -Path $paths.emf -PublishFormat 'pfEMF' -Overwrite ([System.Convert]::ToBoolean($Config.Overwrite))
            }
            if ($null -ne $paths.htm) {
                Write-Host "Exporting Page: " $paths.htm -BackgroundColor Green
                Invoke-OneNotePublish -ID ($page.Id) -Path $paths.emf -PublishFormat 'pfHTML' -Overwrite ([System.Convert]::ToBoolean($Config.Overwrite))
            }

        }
    }
    catch {
        Write-Host $global:error -ForegroundColor Red
        Exit
    }
}

function Get-OneNotePagePublishPath {
    <#
        Helper for publish: returns an object with a certain path set
    #>
    param(
        [string]$ExportFormat,
        [object]$Paths
    )
    try {
        $ExportFormat = $ExportFormat.Trim()
        $Extension = $ExportFormat
        $Dir = $ExportFormat
        if ($ExportFormat -eq "markdown")
        {
            $Extension = "md"
            $Dir = "markdown"
        }
        $path = (Join-Path -Path $Paths.ExportRootPath -ChildPath $Dir | Join-Path -ChildPath $Paths.RelativeRoot | Join-Path -ChildPath $Paths.FullName) + "." + $Extension
        $Paths | Add-Member -Type NoteProperty -Name $ExportFormat -Value $path -Force
        New-Dir -Path ([IO.Path]::GetDirectoryName($path)) | Out-Null
        return $Paths
    }
    catch {
        Write-Host $global:error -ForegroundColor Red
        Exit
    }
}

function Get-OneNotePagePublishPaths {
    <#
        Helper Object to return paths for all export types, given a page as input to publish
        https://docs.microsoft.com/en-us/office/client-developer/onenote/enumerations-onenote-developer-reference#odc_PublishFormat
    #>
    param(
        [Object]$Config,
        [String]$Path,
        [Object]$Page
    )
    try {
        $paths = New-Object -TypeName PSObject
        $paths | Add-Member -Type NoteProperty -Name 'ExportRootPath' -Value ($Config.ExportRootPath) -Force
        $paths | Add-Member -Type NoteProperty -Name 'RelativeRoot' -Value $Path -Force
        $paths | Add-Member -Type NoteProperty -Name 'FullName' -Value $Page.FullName -Force
        $paths | Add-Member -Type NoteProperty -Name 'PageName' -Value $Page.Name  -Force
        $paths | Add-Member -Type NoteProperty -Name 'PageId' -Value $Page.Id  -Force
        # for all files types create helper"
        $exportFormats = $Config.ExportFormat -split ","
        # for each of the export objects specified:
        foreach($exportFormat in $exportFormats) {
            $paths = Get-OneNotePagePublishPath -ExportRootPath $Config.ExportRootPath -ExportFormat $exportFormat -Paths $paths -PageId ($Page.Id)
        }
        return $paths
    }
    catch {
        Write-Host $global:error -ForegroundColor Red
        Exit
    }
}