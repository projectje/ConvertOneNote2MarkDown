function Get-PandocExecutable {
    <#
        Fetches the specified version of Pandoc into temp location and returns location of exec path
        ref: https://jonlabelle.com/snippets/view/powershell/transform-markdown-files-with-pandoc-and-powershell
    #>
    $global:ProgressPreference = 'SilentlyContinue'
    if ($ShowProgress) {$ProgressPreference = 'Continue'}
    $panDocVersion = "2.11.3.2"
    $pandocSourceURL = "https://github.com/jgm/pandoc/releases/download/$panDocVersion/pandoc-$panDocVersion-windows-x86_64.zip"
    $pandocDestinationPath = New-Item (Join-Path ([System.IO.Path]::GetTempPath()) "PanDoc") -ItemType Directory -Force
    $pandocZipPath = Join-Path $pandocDestinationPath "pandoc-$panDocVersion-windows-x86_64.zip"
    $pandocExePath = Join-Path (Join-Path $pandocDestinationPath "pandoc-$panDocVersion") "pandoc.exe"
    if (-not (Test-Path -Path $pandocExePath)) {
        Invoke-WebRequest -Uri $pandocSourceURL -OutFile $pandocZipPath
        Expand-Archive -Path (Join-Path $pandocDestinationPath "pandoc-$panDocVersion-windows-x86_64.zip") -DestinationPath $pandocDestinationPath -Force
    }
    return $pandocExePath
}

function Invoke-ConvertDocxToMd {
    <#
        .SYNOPSIS
            Pandoc Conversion:
            see: https://github.com/jgm/pandoc
            see: https://pandoc.org/MANUAL.html for options
            see: https://gist.github.com/heardk/ded40b72056cee33abb18f3724e0a580
            see: https://github.com/pandoc
            see: https://github.com/jpogran/psdoc
            see: https://jonlabelle.com/snippets/view/powershell/transform-markdown-files-with-pandoc-and-powershell
    #>
    # list of Pandoc options
    param(
        [string] $InputFile, # file
        [string] $InputFormat = "docx", # -f
        [string] $OutputFormat = "markdown", # -t

        [string] $OutputFile, # -o
        [string] $MediaPath                             # --extra-media
    )
    $OutputFormat = "$OutputFormat" + "-simple_tables" + "-multiline_tables" + "-grid_tables+pipe_tables"

    # --wrap=none ensures that text in the new .md files doesn't get wrapped to new lines after 80 characters
    # --atx-headers makes headers in the new .md files appear as # h1, ## h2 and so on

    try {
        pandoc.exe $InputFile -f $InputFormat -t $OutputFormat -o $OutputFile --wrap=none --atx-headers --extract-media=$MediaPath
    }
    catch {
        Write-Host "Error while converting file '$OutputFile' to md: $($Error[0].ToString())" -ForegroundColor Red
    }
}