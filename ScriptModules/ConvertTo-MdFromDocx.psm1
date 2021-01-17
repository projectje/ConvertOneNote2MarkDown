<#
    .SYNOPSIS
        Converts a docx to a desired md formats

        ScriptModule: parent is responsible for including necessary modules
#>
function ConvertTo-MdFromDocx {
    <#
        .SYNOPSIS
            Handle Pandoc Conversion of a published OneNote docx file
    #>
    param (
        [PSCustomObject]$Config,
        [PSCustomObject]$EnrichedPageObject
    )
    try {
        # once docx has been created handle pandoc MD conversions
        # todo: if specifying multiple md formats media will be exported multiple times
        foreach ($exportFormat in $Config.ExportFormat) {
            if ($Config.PandocMdFormats -contains $exportFormat) {
                $documentInputPath = $EnrichedPageObject | Select-Object -ExpandProperty 'docx'
                $documentOutputPath = $EnrichedPageObject | Select-Object -ExpandProperty $exportFormat
                New-Dir -Path  ([IO.Path]::GetDirectoryName($documentOutputPath)) | Out-Null
                Invoke-ConvertDocxToMd -PandocExec $Config.Pandoc -OutputFormat $exportFormat -Inputfile $documentInputPath -OutputFile $documentOutputPath -MediaPath $EnrichedPageObject.MediaPath

                # filters:
                if ($Config.MdClearSpaces -eq $true) {
                    #    Invoke-MdClearSpaces -MdPath $documentOutputPath
                }
                if ($Config.MdClearEspace -eq $true) {
                    #    Invoke-MdClearEscape -MdPath $documentOutputPath
                }
                if ($Config.MdAddYaml -eq $true) {
                    #    Invoke-MdAddYaml -MdPath $documentOutputPath -PageName $Page.Name -PageDateTime $Page.DateTime
                }
                Invoke-MdRenameImages -MdPath $documentOutputPath -MediaPath $EnrichedPageObject.MediaPath -PageName $EnrichedPageObject.Name
                #Invoke-MdImagePathReference -MdPath $documentOutputPath -MediaPath $Page.MediaPath -LevelsPrefix $Level
            }
        }
    }
    catch {
        throw
    }
}