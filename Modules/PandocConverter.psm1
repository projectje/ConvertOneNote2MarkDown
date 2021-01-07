function Invoke-ConvertDocxToMd {
    <#
        .SYNOPSIS
            Convert Word document to MD document
            see: https://pandoc.org/MANUAL.html for options
            see: https://gist.github.com/heardk/ded40b72056cee33abb18f3724e0a580
    #>

    param(
        [string] $InputFormat = "docx", # -f
        [string] $OutputFormat = "markdown", # -t
        [string] $InputFile, # file
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