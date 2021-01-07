
function Remove-InvalidFileNameChars {
    <#
        .SYNOPSIS
            remove invalid characters from a filename
    #>
    param(
        [Parameter(Mandatory = $true,
            Position = 0,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true)]
        [string]$Name
    )
    $newName = $Name.Split([IO.Path]::GetInvalidFileNameChars()) -join '-'
    return (((($newName -replace "\s", "-") -replace "\[", "(") -replace "\]", ")").Substring(0, $(@{$true = 130; $false = $newName.length}[$newName.length -gt 150])))
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
    $rePattern = ($SpecialChars.ToCharArray() |ForEach-Object { [regex]::Escape($_) }) -join "|"
    $newName = $Name.Split([IO.Path]::GetInvalidFileNameChars()) -join '-'
    return ($newName -replace $rePattern, "" -replace "\s", "-")
}

function New-Dir {
    <#
        Creates a dir if not exist
        https://stackoverflow.com/questions/16906170/create-directory-if-it-does-not-exist
    #>
    param(
        [string]$Path
    )
    New-Item -ItemType Directory -Force -Path $Path | Out-Null
}

function Remove-File {
    <#
        Remove a file
    #>
    param (
        [string]$File
    )
    try {
        Remove-Item -path $File -Force -ErrorAction SilentlyContinue
    }
    catch {
        Write-Host $Error -ForegroundColor Red
        Exit
    }
}

function ReplaceStringInFile {
    <#
        Replace string in file
    #>
    param (
        [string]$File,
        [string]$StringToBeReplaced = "",
        [string]$StringThatWillReplaceIt = ""
    )
    ((Get-Content -LiteralPath $File -Raw -encoding utf8).Replace($StringToBeReplaced, '')) | Set-Content -LiteralPath $File
}