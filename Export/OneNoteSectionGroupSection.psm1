<#
    Child Module  (parents needs to reference the dependent modules)
#>
function Export-OneNoteSectionGroupSection {
    <#
        Exports One item in a SectionGroup Collection which can be a sectiongroup collection or a section collection
    #>
    param (
        [Object]$Config,
        [System.Xml.XmlElement]$SectionGroup,
        [int]$Level,
        [string]$Typez,
        [string]$Path

    )
    try {
        $sectionGroupName = $SectionGroup.name | Remove-InvalidFileNameChars
        $dir = Join-Path -Path $Path -ChildPath $sectionGroupName
        if ($Typez -eq "Section") {
            Export-OneNoteSectionCollection -Config $Config -SectionCollection $SectionGroup.Section -Path $dir -Level ($Level + 1)
        }
        elseif ($Typez -eq "SectionGroup") {
            Export-OneNoteSectionGroupCollection -Config $Config -SectionGroupCollection $SectionGroup.SectionGroup -Path $dir -Level ($Level + 1)
        }
    }
    catch
    {
        Write-Host $global:error -ForegroundColor Red
        Exit
    }
}