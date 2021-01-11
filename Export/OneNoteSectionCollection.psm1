<#
    Child Module (parents needs to reference the dependent modules)
#>
function Export-OneNoteSectionCollection {
    <#
        Export Collection of Sections (each holding a page collection)
    #>
    param(
        [Object]$Config,
        [String]$RelativePath,
        [System.Array] $SectionCollection,
        [Int]$Level = 0
    )
    try {
        foreach ($section in $SectionCollection) {
            $sectionFileName = $section.Name | Remove-InvalidFileNameChars
            $NewRelativePath = Join-Path $RelativePath -ChildPath $sectionFileName
            Write-Host "Exporting Section $NewRelativePath" -ForegroundColor Green
            $pageCollection = Get-OneNoteSectionPageCollection -Section $section
            Export-OneNotePageCollection -Config $Config -PageCollection $pageCollection -Level $Level -RelativePath $NewRelativePath
        }
    }
    catch {
        Write-Host $global:error -ForegroundColor Red
        Exit
    }
}