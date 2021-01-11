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
        $Config | Add-Member -Type NoteProperty -Name 'Notebook' -Value $notebookFileName -Force
        $sectionCollection = Get-OneNoteNotebookSectionCollection -Notebook $Notebook
        $sectionGroupCollection = Get-OneNoteNotebookSectionGroupCollection -Notebook $Notebook -Include_Recyclebin $false
        if ($null -ne $sectionCollection) {
            Export-OneNoteSectionCollection -Config $Config -SectionCollection $sectionCollection -RelativePath $notebookFileName
        }
        if ($null -ne $sectionGroupCollection) {
            Export-OneNoteSectionGroupCollection -Config $Config -SectionGroupCollection $sectionGroupCollection -RelativePath $notebookFileName
        }
    }
    catch {
        Write-Host $global:error -ForegroundColor Red
        Exit
    }
}