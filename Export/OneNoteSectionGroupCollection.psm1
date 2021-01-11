<#
    Child Module (parents needs to reference the dependent modules)
#>
function Export-OneNoteSectionGroupCollection {
    <#
        Export SectionGroup Collection. A section group can contain an entry "Section" or an entry "SectionGroup"
    #>
    param (
        [Object]$Config,
        [String]$RelativePath,
        [System.Array]$SectionGroupCollection,
        [Int]$Level = 0
    )
    try {
        if ($SectionGroupCollection.Count -gt 0 -and $null -ne $SectionGroupcollection) {
            Write-Host "Exporting New SectionGroup $RelativePath" -ForegroundColor Green
            $sectionGroupItemsWithSection = Get-OneNoteSectionGroupCollectionSectionCollection -SectionGroup $SectionGroupCollection
            Write-Host "Parsing Entries Section in SectionGroup" -ForegroundColor Green
            foreach ($sectionGroup in $sectionGroupItemsWithSection) {
                Export-OneNoteSectionGroupSection -Config $Config -SectionGroup $sectionGroup -Level $Level -Typez "Section" -RelativePath $RelativePath
            }
            # get all entries that contain an entry "sectiongroup"
            Write-Host "Parsing Entries SectionGroup in SectionGroup" -ForegroundColor Green
            $sectionGroupItemsWithSectionGroup = Get-OneNoteSectionGroupCollectionSectionGroupCollection -SectionGroup $sectionGroupCollection
            foreach ($sectionGroup in $sectionGroupItemsWithSectionGroup ) {
                Export-OneNoteSectionGroupSection -Config $Config -SectionGroup $sectionGroup -Level $Level -Typez "SectionGroup" -RelativePath $RelativePath
            }
        }
    }
    catch {
        Write-Host $global:error -ForegroundColor Red
        Exit
    }
}