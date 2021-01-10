<#
    Child Module (parents needs to reference the dependent modules)
#>
function Export-OneNoteSectionCollection {
    <#
        Export Collection of Sections (each holding a page collection)
    #>
    param(
        [Object]$Config,
        [System.Array] $SectionCollection,
        [String]$Path,
        [Int]$Level = 0
    )
    try {
        foreach ($section in $SectionCollection) {
            $sectionFileName = $section.Name | Remove-InvalidFileNameChars
            $dir = Join-Path $Path -ChildPath $sectionFileName
            Write-Host "Exporting Section $dir" -ForegroundColor Green
            $pageCollection = Get-OneNoteSectionPageCollection -Section $section
            Export-OneNotePageCollection -PageCollection $pageCollection -Path $dir -Level $Level -Config $Config
        }
    }
    catch {
        Write-Host $global:error -ForegroundColor Red
        Exit
    }
}