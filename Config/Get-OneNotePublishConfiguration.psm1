<#
    .SYNOPSIS
        Validates and enriches the export.cfg configuration file > PSCustomObject
#>
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

function Get-OneNotePublishConfiguration {
    <#
        .SYNOPSIS
            Handle Configuration for Export
    #>
    param(
        [string]$Path
    )
    try {
        $config = Get-Config -path $Path
        $config = Export-OneNoteConfigCheckKeyValue -Config $config -Key ExportRootPath -Type "Path" -Required $true |
        Export-OneNoteConfigCheckKeyValue -Config $config -Key AutoDownloadPandoc -Type "Bool"  -Required $false -DefaultValue "True" |
        Export-OneNoteConfigCheckKeyValue -Config $config -Key ExportFormat -Type "Array"  -Required $true |
        Export-OneNoteConfigCheckKeyValue -Config $config -Key Overwrite -Type "Bool"  -Required $false  -DefaultValue "False" |
        Export-OneNoteConfigCheckKeyValue -Config $config -Key OverwriteAttachments -Type "Bool" -Required $false -DefaultValue "False" |
        Export-OneNoteConfigCheckKeyValue -Config $config -Key MdCentralMediaPath -Type "Bool" -Required $false -DefaultValue "False" |
        Export-OneNoteConfigCheckKeyValue -Config $config -Key MdClearSpaces -Type "Bool" -Required $false -DefaultValue "False" |
        Export-OneNoteConfigCheckKeyValue -Config $config -Key MdClearEscape -Type "Bool" -Required $false -DefaultValue "False" |
        Export-OneNoteConfigCheckKeyValue -Config $config -Key MdAddYaml -Type "Bool" -Required $false -DefaultValue "False"

        if ($null -ne $config)
        {
            $TagPath = Join-Path -Path $config.ExportRootPath -ChildPath "_tags"
            New-Dir -Path $TagPath
            # delete the tag items on each run since otherwise they will be appended to
            (Get-ChildItem -Path $config.ExportRootPath -Recurse -Filter '*.tag.html').Fullname | Remove-Item

            $config | Add-Member -Type NoteProperty -Name PublishFormats -Value (Get-OneNotePublishFormats) -Force -PassThru |
                      Add-Member -Type NoteProperty -Name pandocMdFormats -Value (Get-PandocMDOutputFormats) -Force -PassThru |
                      Add-Member -Type NoteProperty -Name TagPath -Value $TagPath -Force

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