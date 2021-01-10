<#
    Child Module (parents needs to reference the dependent modules)
#>

function Export-OneNoteNotebookCollection {
    <#
        Export Collection of all Notebooks
    #>
    param (
        [Object]$Config,
        [System.Xml.XmlElement]$NotebookCollection
    )
    try {
        $count = Get-OneNoteNotebookCollectionCount -NotebookCollection $NotebookCollection
        if ($count -eq 0) {
            Write-Host "Warning: no notebooks found"
        }
        for ($item=0; $item -lt $count; $item++) {
            Write-Host "Exporting Notebook $($item+1) out of $count" -ForegroundColor Green
            $notebook = Get-OneNoteNotebook -NotebookCollection $NotebookCollection -NotebookItem $item
            Write-Host "Exporting Notebook:" -ForegroundColor Green
            $notebook
            Export-OneNoteNoteBook -Notebook $notebook -Config $Config
        }
        Write-Host "*DONE*" -ForegroundColor Green
    }
    catch {
        Write-Host $global:error -ForegroundColor Red
        Exit
    }
}