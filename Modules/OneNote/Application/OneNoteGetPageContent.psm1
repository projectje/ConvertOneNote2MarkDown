<#
    Operations using OneNote.Application.GetContents
#>
function Get-OneNotePageInsertedFileObjects {
    <#
        export file objects in a page and write
    #>
    param(
        [string]$Id,
        [string]$AttachmentsPath,
        [bool]$OverwriteAttachments
    )
    try {
        $OneNoteContent = New-Object -ComObject OneNote.Application
        [xml]$pagexml = $null
        $OneNoteContent.GetPageContent($Id, [ref]$pagexml, 7)
        $insertedFileCollection = $pagexml.Page.Outline.OEChildren.OE | Where-Object { $_.InsertedFile }
        $fileArray = New-Object -TypeName "System.Collections.ArrayList"
        foreach ($insertedFile in $insertedFileCollection) {
            $name = $insertedFile.InsertedFile.preferredName | Remove-InvalidFileNameCharsInsertedFiles
            $destination = Join-Path -Path $AttachmentsPath -ChildPath $name
            $destinationExists = Test-Path -Path $destination
            if (($true -eq $destinationExists -and $Config.OverwriteAttachments -eq $true) -or ($false -eq $destinationExists)) {
                New-Item -Path $AttachmentsPath -ItemType "directory" -Force | Out-Null
                # not completely safe but works for me:
                if ($destination.Length -gt 259) {
                    $extension = [System.IO.Path]::GetExtension($destination)
                    $destination = $destination.Substring(0,250) + $extension
                }
                Write-Host "Publishing Attachment: " $destination -ForegroundColor Blue
                Copy-Item -Path "$($insertedFile.InsertedFile.pathCache)" -Destination $destination -Force
                $fileArray.Add($name) | Out-Null
            } elseif ($false -eq (Test-Path -Path $destination)) {
                Write-Host "Skipping Attachment: " $destination -ForegroundColor Yellow
            }
        }
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($OneNoteContent) | Out-Null
        Remove-Variable OneNoteContent
        return $fileArray
        # todo return filename array
        # todo for md copy the export from docx since no use doing this twice
        # todo will crash on very long names (> 259 characters)
    }
    catch {
        throw
    }
}

function Get-OneNotePageXML {
    <#
        return page xml
    #>
    param(
        [string]$Id
    )
    try {
        $OneNoteContent = New-Object -ComObject OneNote.Application
        [xml]$pagexml = $null
        $OneNoteContent.GetPageContent($Id, [ref]$pagexml, 7)
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($OneNoteContent) | Out-Null
        Remove-Variable OneNoteContent
        return $pagexml.Page
    }
    catch {
        throw
    }
}

# {0DDEA6C2-B044-452B-A5DB-143CB495A52C}{33}{B0}

function Get-OneNoteBinaryPageContent {
    <#
        return page xml
    #>
    param(
        [string]$PageId,
        [string]$ObjectId
    )
    try {
        $OneNoteContent = New-Object -ComObject OneNote.Application
        [string]$binaryContent = $null
        $OneNoteContent.GetBinaryPageContent($PageId, $ObjectId, [ref]$binaryContent)
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($OneNoteContent) | Out-Null
        Remove-Variable OneNoteContent
        return $binaryContent
    }
    catch {
        throw
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
        throw
    }
}