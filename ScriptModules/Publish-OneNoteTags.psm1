<#
    .SYNOPSIS
        publishes complete onenotehierarchy in pages, attachments, etc via configuration in configuration file

        ScriptModule: parent is responsible for including necessary modules
#>
function Publish-OneNoteTags {
    param (
        [PSCustomObject]$Config,
        [hashtable]$TagsTable
    )
    $names = $TagsTable.GetEnumerator().Name
    foreach($key in $names) {
        $output = '<html><head></head><body><ul>'
        $values = $TagsTable[$key]
        $tagPath = Join-Path -Path $Config.TagPath -ChildPath ($key + '.tag.html')
        foreach($value in $values) {
            $value = $value -replace '<[^>]+>',''
            $output += "<li>" + $value + "</li>`r`n"
        }
        $output = $output + "</ul></body></html>"
        Add-Content $tagPath $output
    }
}