Import-Module "$PSScriptRoot\OneNotePage.psm1" -Force
<#
    Handles operations on Page Collection
#>
function Get-OneNotePageCollectionCount {
    <#
        Gets the count of pages in a page collection
    #>
    param(
        [System.Array]$PageCollection
    )
    return $PageCollection.Count
}

function Get-OneNotePageFromPageCollection {
    <#
        Gets a page by ID from the page collection
    #>
    param(
        [System.Array]$PageCollection,
        [string]$ID
    )
    return $PageCollection | Where-Object { $_.ID -eq $ID}
}

function Get-OneNotePageHasChildren {
    <#
        Returns true if the page has children
    #>
    param(
        [System.Array]$PageCollection,
        [string]$ID
    )
    # for loop in page collection until we hit the ID, then check if the next item has a level higher than the current one
    $count = Get-OneNotePageCollectionCount -PageCollection $PageCollection
    for($i=0;$i -lt $count; $i++) {
        if ($PageCollection[$i].ID -eq $ID) {
            $currentPageLevel = $PageCollection[$i].pageLevel
            if ($PageCollection[$i+1].pageLevel -gt $currentPageLevel) {
                return $true;
            }
        }
    }
    return $false;
}

function Get-OneNoteEnrichPageCollection {
    <#
        Adds some extra properties to a page in page collection useful during conversion / export
    #>
    param(
        [System.Array]$PageCollection
    )
    $count = Get-OneNotePageCollectionCount -PageCollection $pageCollection
    $Dir = "empty"
    $SubDir = "empty"

    for($i=0;$i -lt $count; $i++) {
        $page = $pageCollection[$i]

        # Add property to indicate if a page has children
        $hasChildren = Get-OneNotePageHasChildren -PageCollection $pageCollection -ID ($page.ID)
        if ($hasChildren) {
            Add-Member -InputObject $page -MemberType NoteProperty -Name "HasChildren" -Value $true -Force
        } else {
            Add-Member -InputObject $page -MemberType NoteProperty -Name "HasChildren" -Value $false -Force
        }

        # Add path to indicate the path
        $path = ""
        if ($page.pageLevel -eq 3) {
            $path = Join-Path $Dir -ChildPath $SubDir
        }
        elseif ($page.pageLevel -eq 2) {
            if ($page.HasChildren) {
                $SubDir = $page.name
                $path = Join-Path $Dir -ChildPath $SubDir
            }
            else
            {
                $path = $Dir
            }
        }
        elseif ($page.pageLevel -eq 1) {
            if ($page.HasChildren) {
                $Dir = $page.name
                $path = $Dir
            }
            else
            {
                $path = ""
            }
        }
        Add-Member -InputObject $page -MemberType NoteProperty -Name "Path" -Value $path -Force

        # change the name in a filtered name
        $page.name = Get-OneNotePageName -Page $page

        $PageCollection[$i] = $page

    }
    return $PageCollection
}