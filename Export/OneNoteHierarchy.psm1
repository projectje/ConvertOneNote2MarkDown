<#
    Child Module (parents needs to reference the dependent modules)
#>
function Export-OneNoteHierarchy {
    <#
        Exports OneNote Hierarchy of all notebooks
    #>
    param (
        [Object]$Config
    )
    try {
        $hierarchy = Get-OneNoteHierarchy
        $notebookCollection = Get-OneNoteNoteBookCollection -Hierarchy $hierarchy
        Export-OneNoteNotebookCollection -Config $Config -NotebookCollection $notebookCollection
    }
    catch {
        Write-Host $global:error -ForegroundColor Red
        Exit
    }
}
