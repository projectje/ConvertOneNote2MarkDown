<#
    .SYNOPSIS
        Tag Handling: returns hasharray with all tags
#>

function Get-OneNoteTags {
    param(
        [PSCustomObject]$EnrichedPageObject,
        [hashtable]$TagsTable
    )
    $xml = Get-OneNotePageXML -Id $EnrichedPageObject.Id

    # check if there are any tags on the page
    $tagNames = @()
    if ($xml.TagDef.Length -gt 0) {
        $countTags = $xml.TagDef.count
        for ($i = 0; $i -lt $countTags; $i++) {
            $tag = [PSCustomObject]@{
                index          = $xml.TagDef[$i].index
                type           = $xml.TagDef[$i].type
                symbol         = $xml.TagDef[$i].symbol
                fontColor      = $xml.TagDef[$i].fontcolor
                highlightColor = $xml.TagDef[$i].highlightcolor
                name           = $xml.TagDef[$i].name
            }
            $tagNames += $tag
        }
    }
    else
    {
        return $TagsTable;
    }

    $countchildren = $xml.Outline.OEChildren.OE.count
    for ($i = 0; $i -lt $countchildren; $i++) {
        if ($null -ne $xml.Outline.OEChildren.OE[$i].Tag) {
            # props
            $task_description = $xml.Outline.OEChildren.OE[$i].T."#cdata-section"
            $task_tempname = $tagNames | Where-Object -Property index -eq $xml.Outline.OEChildren.OE[$i].Tag.index | Select-Object name
            $task_name = $task_tempname.name # this can be null when multiple tags are placed directly behind each other
            $tag_key = $task_name
            if ($xml.Outline.OEChildren.OE[$i].Tag.completed -eq 'true') {$tag_key = $tag_key + "_completed"}

            # add to array
            if ($TagsTable.ContainsKey($tag_key)) {
                $tag_existing_values = $TagsTable[$tag_key]
                $tag_new_values = $tag_existing_values + $task_description
                $TagsTable[$tag_key] = $tag_new_values
            }
            else {
                [array]$tag_new_values = $task_description
                $TagsTable[$tag_key] = $tag_new_values
            }
        }
    }
    return $TagsTable

    # other possible attributes to enrich in this output:
    # $EnrichedPageObject.RelativePath
    # $EnrichedPageObject.FullName
    # $xml.Outline.OEChildren.OE[$i].Tag
    # $xml.Outline.OEChildren.OE[$i].Tag.index
    # $xml.Outline.OEChildren.OE[$i].Tag.completed
    # $xml.Outline.OEChildren.OE[$i].Tag.disabled
    # $xml.Outline.OEChildren.OE[$i].Tag.creationDate
    # $xml.Outline.OEChildren.OE[$i].Tag.completionDate
}