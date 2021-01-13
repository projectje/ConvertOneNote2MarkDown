function Export-OneNoteConfigCheckKeyValue {
    <#
        .SYNOPSIS
            Checks if a config value exists
    #>
    param(
        [object]$Config,
        [string]$Key,
        [string]$Type,
        [bool]$Required,
        [string]$DefaultValue
    )
    try {
        if ($null -eq $Config) { return $null }

        if ($null -eq ($Config.PSobject.Properties.name -match $Key)) {
            if ($Required -eq $true) {
                Write-Host "$Key required" -BackgroundColor -Red
                return $null
            }
            else {
                $Config | Add-Member -Force -Type NoteProperty -Name $Key -Value $DefaultValue
            }
        }

        $Value = $Config | Select-Object -ExpandProperty $Key
        if ($null -eq $Value -or $Value.Trim() -eq "") {
            if ($Required -eq $true) {
                Write-Host "$Value of $Key incorrect or empty" -BackgroundColor -Red
                return $null
            }
            else {
                $Config | Add-Member -Force -Type NoteProperty -Name $Key -Value $DefaultValue
            }
        }

        if ($Type -eq "Path") {
            if ((Test-Path -Path $Value) -ne $true) {
                New-Dir -Path $path
            }
            $Config | Add-Member -Force -Type NoteProperty -Name $Key -Value $Value.Trim()
        }
        elseif ($Type -eq "Bool") {
            [bool]$boolValue = [System.Convert]::ToBoolean($Value.Trim())
            $Config | Add-Member -Force -Type NoteProperty -Name $Key -Value $boolValue
        }
        elseif ($Type -eq "Array") {
            [array]$arrayValue = $Value -split ',' -replace '^\s+|\s+$'
            $config | Add-Member -Force -Type NoteProperty -Name $Key -Value $arrayValue
        }

        return $Config
    }
    catch {
        throw
    }
}

function Export-Config {
    <#
        .SYNOPSIS
            Handle Configuration for Export
    #>
    param(
        [string]$Path
    )
    try {
        $config = Get-Config -path $Path
        $config = Export-OneNoteConfigCheckKeyValue -Config $config -Key ExportRootPath -Type "Path" -Required $true
        $config = Export-OneNoteConfigCheckKeyValue -Config $config -Key AutoDownloadPandoc -Type "Bool"  -Required $false -DefaultValue "True"
        $config = Export-OneNoteConfigCheckKeyValue -Config $config -Key ExportFormat -Type "Array"  -Required $true
        $config = Export-OneNoteConfigCheckKeyValue -Config $config -Key Overwrite -Type "Bool"  -Required $false  -DefaultValue "False"
        $config = Export-OneNoteConfigCheckKeyValue -Config $config -Key OverwriteAttachments -Type "Bool" -Required $false -DefaultValue "False"
        $config = Export-OneNoteConfigCheckKeyValue -Config $config -Key MdCentralMediaPath -Type "Bool" -Required $false -DefaultValue "False"
        $config = Export-OneNoteConfigCheckKeyValue -Config $config -Key MdClearSpaces -Type "Bool" -Required $false -DefaultValue "False"
        $config = Export-OneNoteConfigCheckKeyValue -Config $config -Key MdClearEscape -Type "Bool" -Required $false -DefaultValue "False"
        $config = Export-OneNoteConfigCheckKeyValue -Config $config -Key MdAddYaml -Type "Bool" -Required $false -DefaultValue "False"

        if ($null -ne $config)
        {
            $config | Add-Member -Type NoteProperty -Name 'PublishFormats' -Value (Get-OneNotePublishFormats) -Force
            $config | Add-Member -Type NoteProperty -Name 'pandocMdFormats' -Value (Get-PandocMDOutputFormats) -Force

            if ($config.AutoDownloadPandoc -eq $true) {
                foreach($exportFormat in $config.ExportFormat) {
                    if (-not ($config.PublishFormats -contains $exportFormat)) {
                        $config | Add-Member -Type NoteProperty -Name 'Pandoc' -Value (Get-PandocExecutable) -Force
                    }
                }
            }

            foreach($exportFormat in $config.ExportFormat) {
                if ($config.PandocMdFormats -contains $exportFormat) {
                    # make sure docx is also selected
                    if (-not ($config.ExportFormat -contains "docx")) {
                        $config.ExportFormat += "docx"
                        $config | Add-Member -Force -Type NoteProperty -Name ExportFormat -Value $config.ExportFormat
                    }
                }
            }
        }

        return $config
    }
    catch {
        throw
    }
}