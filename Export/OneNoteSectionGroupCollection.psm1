<#
    Child Module (parents needs to reference the dependent modules)
#>
function Export-OneNoteSectionGroupCollection {
    <#
        Export SectionGroup Collection. A section group can contain an entry "Section" or an entry "SectionGroup"
    #>
    param (
        [Object]$Config,
        [System.Array]$SectionGroupCollection,
        [String]$Path,
        [Int]$Level = 0
    )
    try {
        if ($SectionGroupCollection.Count -gt 0 -and $null -ne $SectionGroupcollection) {
            Write-Host "Exporting New SectionGroup $Path" -ForegroundColor Green
            $sectionGroupItemsWithSection = Get-OneNoteSectionGroupCollectionSectionCollection -SectionGroup $SectionGroupCollection
            Write-Host "Parsing Entries Section in SectionGroup" -ForegroundColor Green
            foreach ($sectionGroup in $sectionGroupItemsWithSection) {
                Export-OneNoteSectionGroupSection -Config $Config -SectionGroup $sectionGroup -Path $Path -Level $Level -Typez "Section"
            }
            # get all entries that contain an entry "sectiongroup"
            Write-Host "Parsing Entries SectionGroup in SectionGroup" -ForegroundColor Green
            $sectionGroupItemsWithSectionGroup = Get-OneNoteSectionGroupCollectionSectionGroupCollection -SectionGroup $sectionGroupCollection
            foreach ($sectionGroup in $sectionGroupItemsWithSectionGroup ) {
                Export-OneNoteSectionGroupSection -Config $Config -SectionGroup $sectionGroup -Path $Path -Level $Level -Typez "SectionGroup"
            }
        }
    }
    catch {
        Write-Host $global:error -ForegroundColor Red
        Exit
    }
}