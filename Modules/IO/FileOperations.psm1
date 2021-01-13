<#
    .SYNOPSIS
        Helper functions for file operations
#>
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
    try {
        $newName = $Name.Split([IO.Path]::GetInvalidFileNameChars()) -join '-'
        return (((($newName -replace "\s", "-") -replace "\[", "(") -replace "\]", ")").Substring(0, $(@{$true = 130; $false = $newName.length}[$newName.length -gt 150])))
    }
    catch
    {
        throw
    }
}

function New-Dir {
    <#
        Creates a dir if not exist
        https://stackoverflow.com/questions/16906170/create-directory-if-it-does-not-exist
    #>
    param(
        [string]$Path
    )
    try {
        New-Item -ItemType Directory -Force -Path $Path | Out-Null
    }
    catch {
        throw
    }
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
        throw
    }
}
