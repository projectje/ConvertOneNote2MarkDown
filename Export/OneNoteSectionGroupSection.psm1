<#
    Child Module  (parents needs to reference the dependent modules)
#>
function Export-OneNoteSectionGroupSection {
    <#
        Exports One item in a SectionGroup Collection which can be a sectiongroup collection or a section collection
    #>
    param (
        [Object]$Config,
        [String]$RelativePath,
        [System.Xml.XmlElement]$SectionGroup,
        [int]$Level,
        [string]$Typez
    )
    try {
        $sectionGroupName = $SectionGroup.name | Remove-InvalidFileNameChars
        $NewRelativePath = Join-Path $RelativePath -ChildPath $sectionGroupName
        if ($Typez -eq "Section") {
            Export-OneNoteSectionCollection -Config $Config -SectionCollection $SectionGroup.Section -Level ($Level + 1) -RelativePath $NewRelativePath
        }
        elseif ($Typez -eq "SectionGroup") {
            Export-OneNoteSectionGroupCollection -Config $Config -SectionGroupCollection $SectionGroup.SectionGroup -Level ($Level + 1) -RelativePath $NewRelativePath
        }
    }
    catch
    {
        Write-Host $global:error -ForegroundColor Red
        Exit
    }
}