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

function Invoke-Pandoc {
    <#
        .SYNOPSIS
            Calls Pandoc with various options (see https://pandoc.org/MANUAL.html for options)

            Todo: this will be a more generic call to pandoc for more formats and options
    #>
    param(
        [string] $InputFile,        #
        [string] $OutputFile,       # -o                        outputfile
        [string] $StandaloneFile,   # -s / --standalone         by default pandoc produces a fragment
        [string] $FileScope,         # --file-scope             by default pandoc will concatenate multiple passed files
        [string] $InputFormat,       # -f / --from              the format to convert from
        [string] $OutputFormat      # =t / --to                 the format to convert to
    )
}

function Invoke-ConvertDocxToMd {
    <#
        .SYNOPSIS
            Pandoc Conversion that was based on https://gist.github.com/heardk/ded40b72056cee33abb18f3724e0a580
    #>
    # list of Pandoc options
    param(
        [string] $PandocExec,
        [string] $InputFile,                        # file
        [string] $InputFormat = "docx",             # -f
        [string] $OutputFormat = "markdown",        # -t
        [string] $OutputFile,                       # -o
        [string] $MediaPath                         # --extra-media
    )

    #
    # todo: not all formats support all extensions e.g. gfm does not support simple table, so should be part of configuration file > handier
    #

    $pandocArgs = @(
        $InputFile,
        "--from=$InputFormat",
        "--to=$($OutputFormat)-simple_tables-multiline_tables-grid_tables+pipe_tables",
        "--output=$OutputFile",
        "--wrap=none", # ensures that text in the new .md files doesn't get wrapped to new lines after 80 characters
        "--markdown-headings=atx", # makes headers in the new .md files appear as # h1, ## h2 and so on
        "--extract-media=$MediaPath"
    )

    try {
        & $PandocExec $pandocArgs
        Write-Host "Publishing Page: " $OutputFile -ForegroundColor Green
    }
    catch {
        Write-Host $global:error -ForegroundColor Red
        Exit
    }
}

function Get-PandocMDOutputFormats {
    $publishformats = @('markdown', 'markdown_mmd', 'markdown_phpextra', 'markdown_strict', 'commonmark', 'commonmark_x', 'gfm', 'markdown_github')
    return $publishformats
}