<#
    .SYNOPSIS
        Operations concerning OneNote.Application.Publish which is also worded as "export"
#>

function Invoke-OneNotePublish {
    <#
        .SYNOPSIS
            Publishes a OneNote Object
            See: https://docs.microsoft.com/en-us/office/client-developer/onenote/enumerations-onenote-developer-reference#odc_PublishFormat
            for publish options (mht, pdf, xps, doc, emf, html)

            Syntax
	            HRESULT Publish(
                    [in]BSTR bstrHierarchyID,
                    [in]BSTR bstrTargetFilePath,
                    [in,defaultvalue(pfOneNote)]PublishFormat pfPublishFormat,
                    [in,defaultvalue(0)]BSTR bstrCLSIDofExporter);
    #>
    param (
        [string] $Id, # [in]BSTR bstrHierarchyID
        [string] $Path, # [in]BSTR bstrTargetFilePath
        [string] $PublishFormat = 'pfWord', # [in,defaultvalue(pfOneNote)]PublishFormat pfPublishFormat,
        [bool] $Overwrite = $true
    )

    if ($PublishFormat -eq 'docx')
    {
        $PublishFormat = 'pfWord'
    }
    elseif ($PublishFormat -eq 'doc')
    {
        $PublishFormat = 'pfWord'
    }
    elseif ($PublishFormat -eq 'one') {
        $PublishFormat = 'pfOneNote'
    }
    elseif ($PublishFormat -eq 'onenote') {
        $PublishFormat = 'pfOneNote'
    }
    elseif ($PublishFormat -eq 'onepkg') {
        $PublishFormat = 'pfOneNotePackage'
    }
    elseif ($PublishFormat -eq 'mht') {
        $PublishFormat = 'pfMHTML'
    }
    elseif ($PublishFormat -eq 'mhtml') {
        $PublishFormat = 'pfMHTML'
    }
    elseif ($PublishFormat -eq 'pdf') {
        $PublishFormat = 'pfPDF'
    }
    elseif ($PublishFormat -eq 'xps') {
        $PublishFormat = 'pfXPS'
    }
    elseif ($PublishFormat -eq 'emf') {
        $PublishFormat = 'pfEMF'
    }
    elseif ($PublishFormat -eq 'htm') {
        $PublishFormat = 'pfHTML'
    }
    else {
        return
    }

    try {
        $OneNotePage = New-Object -ComObject OneNote.Application
        # $OneNotePage | getMember
        [bool] $fileExists = [System.IO.File]::Exists($Path)
        if ($fileExists -eq $true -and $Overwrite -eq $true) {
            Remove-Item -path $Path -Force -ErrorAction SilentlyContinue
            $fileExists = $false
        }
        if ($fileExists -eq $false) {
            Write-Host "Publishing Page: " $Path -ForegroundColor Green
            $dirPath = [IO.Path]::GetDirectoryName($Path)
            New-Item -Path $dirPath -ItemType "directory" -Force | Out-Null
            $OneNotePage.Publish($Id, $Path, $PublishFormat, "")
        }
        else {
            Write-Host "Skipping Page: " $Path -ForegroundColor Yellow
        }
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($OneNotePage) | Out-Null
        Remove-Variable OneNotePage
    }
    catch {
        write-host $Id $Path $PublishFormat
        Throw
    }
}

function Get-OneNotePublishFormats {
    $publishformats = @('doc', 'docx', 'pdf', 'xml', 'one', 'onepkg', 'htm', 'mht', 'emf')
    return $publishformats
}


function Get-OneNotePageHasChildren {
    <#
        Returns true if the page has children
    #>
    param(
        [System.Array]$PageCollection,
        [string]$ID
    )
    try {
        # for loop in page collection until we hit the ID, then check if the next item has a level higher than the current one
        $count = $PageCollection.Count
        for ($i = 0; $i -lt $count; $i++) {
            if ($PageCollection[$i].ID -eq $ID) {
                $currentPageLevel = $PageCollection[$i].pageLevel
                if ($PageCollection[$i + 1].pageLevel -gt $currentPageLevel) {
                    return $true;
                }
            }
        }
        return $false;
    }
    catch {
        Throw
    }
}

function Get-OneNoteEnrichPageCollection {
    <#
        Adds some extra properties to a page in page collection useful before a publish
        to make a hierarchy of nested pages

        On the way the page structure itself (3 level) should be structured are different opinions, so
        probably will be a parameter

        Also the possibility exists to auto unfold page structures (otherwise they will not be exported)

    #>
    param(
        [System.Array]$PageCollection,
        [String]$RelativePath
    )
    $count = $PageCollection.Count
    $Dir = "empty"
    $SubDir = "empty"
    $pageArray = New-Object -TypeName "System.Collections.ArrayList"
    $namesArray = New-Object -TypeName "System.Collections.ArrayList"
    try {
        for ($i = 0; $i -lt $count; $i++) {
            $page = $PageCollection[$i]

            # 1. determine if has children (needed by 2)
            $hasChildren      = Get-OneNotePageHasChildren -PageCollection $pageCollection -ID $page.ID;

            # 2. determine pagePath
            $path = ""
            $name = Remove-InvalidFileChars -Name $page.name
            if ($page.pageLevel -eq 3)
            {
                $path = Join-Path -Path $Dir -ChildPath $SubDir
            }
            elseif ($page.pageLevel -eq 2)
            {
                if ($hasChildren) {
                    $SubDir = $name
                    $path = Join-Path -Path $Dir -ChildPath $SubDir
                }
                else {
                    $path = $Dir
                }
            }
            elseif ($page.pageLevel -eq 1) {
                if ($hasChildren) {
                    $Dir = $name
                    $path = $Dir
                }
                else {
                    $path = ""
                }
            }

            # 3. If the collection contains duplicate names, add and extension to one name
            $fullName = ""
            if ($null -ne $path -and "" -ne $path) {
                $fullName = Join-Path -Path $path -ChildPath $name
            }
            else {
                $fullName = $name
            }
            if ($namesArray.Contains($fullName)) {
                $postfix = 0
                $testName = $fullName
                while ($namesArray.Contains($testName)) {
                    $testName = $fullName + "-" + ($postfix++)
                }
                $fullName = $testName
            }
            $namesArray.Add($fullName) | Out-Null

            # Report properties we are not handling
            $page | Get-Member -MemberType Property | ForEach-Object {
                $handledProperties = 'ID', 'lastModifiedTime', 'dateTime', 'pageLevel', 'name', 'isUnread', 'Meta', 'stationeryName', 'isCurrentlyViewed', 'isCollapsed'
                if (-not ($handledProperties -contains $_.Name)) {
                    write-host "Warning, Property not yet handled: " $_.Name
                }
            }

            $pageObject = [PSCustomObject] @{
                Id               = $page.Id
                Name             = $name
                DateTime         = $page.dateTime
                LastModifiedTime = $page.lastModifiedTime
                IsUnread         = $page.isUnread # not used in script
                IsCurrentlyViewed = $page.IsCurrentlyViewed # not used in script
                StationeryName    = $page.stationeryName # not used in script
                Meta              = $page.Meta # not used in script
                IsCollapsed = $page.isCollapsed # not used in script
                PageLevel        = $page.pageLevel
                HasChildren      = $hasChildren
                RelativePath     = $RelativePath
                Path             = $path
                FullName = $fullName
            }

            # I read collapsed pages would not be exported but i dont see this happening with my collapsed pages so commented
            # Otherwise the TreeCollapsedStateType would needed to be called with tcsExpanded
            if ($null -ne $pageObject.IsCollapsed ) {
                #write-host "Warning the page " $pageObject.Name " is collapsed  " -ForegroundColor red
            }

            $pageArray.Add($pageObject) | Out-Null
        }
        return $pageArray
    }
    catch {
        Throw
    }
}

function Remove-InvalidFileChars {
    <#
        .SYNOPSIS
            remove invalid characters from a filename
    #>
    param(
        [Parameter(Mandatory = $true,
            Position = 0,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true)]
        [string]$Name
    )
    try {
        $newName = $Name.Split([IO.Path]::GetInvalidFileNameChars()) -join '-'
        return (((($newName -replace "\s", "-") -replace "\[", "(") -replace "\]", ")").Substring(0, $(@{$true = 130; $false = $newName.length}[$newName.length -gt 150])))
    }
    catch
    {
        throw
    }
}