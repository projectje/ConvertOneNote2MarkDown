<#
    .SYNOPSIS
        filters applied after writing an MD file
        see: https://github.com/nixsee/ConvertOneNote2MarkDown/blob/master/ConvertOneNote2MarkDown-v2.ps1
#>
function Invoke-MdRenameImages {
    <#
        Rename images to have unique names  - NoteName-Image#-HHmmssff.xyz
    #>
    param (
        [string]$MdPath,
        [string]$MediaPath,
        [string]$PageName
    )
    try {
        $timeStamp = (Get-Date -Format HHmmssff).ToString()
        $timeStamp = $timeStamp.replace(':', '')
        $images = Get-ChildItem -Path "$($MediaPath)/media" -Include "*.png", "*.gif", "*.jpg", "*.jpeg" -Recurse -ErrorAction SilentlyContinue | Where-Object { $_.Name.SubString(0, 5) -match "image" }
        foreach ($image in $images) {
            $newimageName = "$($PageName.SubString(0,[math]::min(30,$PageName.length)))-$($image.BaseName)-$($timeStamp)$($image.Extension)"
            Rename-Item -Path "$($image.FullName)" -NewName $newimageName -ErrorAction SilentlyContinue
            ((Get-Content -path $MdPath -Raw).Replace("$($image.Name)", "$($newimageName)")) | Set-Content -Path $MdPath
        }
    }
    catch {
        Write-Host $global:error -ForegroundColor Red
        $global:Error
        Exit
    }
}

function Invoke-MdAddYaml {
    <#
        Add Yaml to top of exported MD
    #>
    Param(
        [string]$MdPath,
        [string]$PageName,
        [string]$PageDateTime
    )
    $orig = Get-Content -path $MdPath
    $orig[0] = "# $PageName"
    $insert1 = $PageDateTime
    $insert1 = [Datetime]::ParseExact($insert1, 'yyyy-MM-ddTHH:mm:ss.fffZ', $null)
    $insert1 = $insert1.ToString("yyyy-MM-dd HH:mm:ss
        ")
    $insert2 = "---"
    Set-Content -Path $MdPath -Value $orig[0..0], $insert1, $insert2, $orig[6..$orig.Length]
}

function Invoke-MdClearSpaces {
    <#
    Clear double spaces from bullets and nonbreaking spaces from blank lines
    #>
    Param (
        [string]$MdPath
    )
    try {
        ((Get-Content -Path $MdPath -Raw -encoding utf8).Replace(">", "").Replace("<", "").Replace([char]0x00A0, [char]0x000A).Replace([char]0x000A, [char]0x000A).Replace("`r`n`r`n", "`r`n")) | Set-Content -Path  $MdPath
    }
    catch {
        Write-Host $global:error -ForegroundColor Red
        $global:Error
        Exit
    }
}

function Invoke-MdClearEscape {
    <#
        Clear Escape
    #>
    Param (
        [string]$MdPath
    )
    Invoke-ReplaceStringInFile -File $MdPath -StringToBeReplaced "\" -StringThatWillReplaceIt ""
}

function Invoke-MdImagePathReference {
    <#
        .SYNOPSIS
        Change MD file Image Path References in Markdown / HTML
    #>
    param (
        [string]$MdPath,
        [string]$MediaPath,
        [string]$LevelsPrefix
    )
    try {

        Write-Host Replace("$($MediaPath.Replace("\","\\"))", "$($LevelsPrefix)")

        #((Get-Content -path $MdPath  -Raw).Replace("$($MediaPath.Replace("\","\\"))", "$($LevelsPrefix)")) | Set-Content -Path $MdPath
        #((Get-Content -path $MdPath  -Raw).Replace("$($MediaPath)", "$($LevelsPrefix)")) | Set-Content -Path $MdPath
    }
    catch
    {
        Write-Host $global:error -ForegroundColor Red
        $global:Error
        Exit
    }
}

function Invoke-ReplaceStringInFile {
    <#
        Replace string in file
    #>
    param (
        [string]$File,
        [string]$StringToBeReplaced = "",
        [string]$StringThatWillReplaceIt = ""
    )
    ((Get-Content -Path $File -Raw -encoding utf8).Replace($StringToBeReplaced, '')) | Set-Content -Path $File
}