<#
    Child Module (parents needs to reference the dependent modules)
#>
function Export-OneNoteNoteBook {
    <#
        Export one OneNote Notebook (which contains sections and sectiongroups)
    #>
    param (
        [Object]$Config,
        [System.Xml.XmlElement]$Notebook
    )
    try {
        Write-Host "Exporting Sections and SectionGroups from Notebook $notebookFileName" -ForegroundColor Green
        $notebookFileName = $Notebook.Name | Remove-InvalidFileNameChars
        $sectionCollection = Get-OneNoteNotebookSectionCollection -Notebook $Notebook
        $sectionGroupCollection = Get-OneNoteNotebookSectionGroupCollection -Notebook $Notebook -Include_Recyclebin $false
        if ($null -ne $sectionCollection) {
            Export-OneNoteSectionCollection -SectionCollection $sectionCollection -Path $notebookFileName -Config $Config
        }
        if ($null -ne $sectionGroupCollection) {
            Export-OneNoteSectionGroupCollection -SectionGroupCollection $sectionGroupCollection -Path $notebookFileName -Config $Config
        }
    }
    catch {
        Write-Host $global:error -ForegroundColor Red
        Exit
    }
}