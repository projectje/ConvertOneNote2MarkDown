<#
    Operations using OneNote.Application.GetContents
#>
function Get-OneNotePageInsertedFileObjects {
    <#
        export file objects in a page and write
    #>
    param(
        [string]$ID,
        [string]$AttachmentsPath
    )
    try {
        $OneNoteContent = New-Object -ComObject OneNote.Application
        [xml]$pagexml = $null
        $OneNoteContent.GetPageContent($Id, [ref]$pagexml, 7)
        $insertedFileCollection = $pagexml.Page.Outline.OEChildren.OE | Where-Object { $_.InsertedFile }
        foreach ($insertedFile in $insertedFileCollection) {
            New-Item -Path $AttachmentsPath -ItemType "directory" -Force | Out-Null
            $destination = Join-Path -Path $AttachmentsPath -ChildPath ($insertedFile.InsertedFile.preferredName | Remove-InvalidFileNameCharsInsertedFiles)
            Write-Host "Publishing Attachment: " $destination -ForegroundColor Blue
            Copy-Item -Path "$($insertedFile.InsertedFile.pathCache)" -Destination $destination -Force
        }
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($OneNoteContent) | Out-Null
        Remove-Variable OneNoteContent
        return $insertedFileCollection
        # todo return filename array
        # todo for md copy the export from docx since no use doing this twice
        # todo will crash on very long names
    }
    catch {
        Write-Host $global:error -ForegroundColor Red
        Exit
    }
}

function Remove-InvalidFileNameCharsInsertedFiles {
    <#
        remove invalid characters from a filename
    #>
    param(
        [Parameter(Mandatory = $true,
            Position = 0,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true)]
        [string]$Name,
        [string]$Replacement = "",
        [string]$SpecialChars = "#$%^*[]'<>!@{};"

    )
    try {
        $rePattern = ($SpecialChars.ToCharArray() |ForEach-Object { [regex]::Escape($_) }) -join "|"
        $newName = $Name.Split([IO.Path]::GetInvalidFileNameChars()) -join '-'
        return ($newName -replace $rePattern, "" -replace "\s", "-")
    }
    catch {
        Write-Host $global:error -ForegroundColor Red
        Exit
    }
}