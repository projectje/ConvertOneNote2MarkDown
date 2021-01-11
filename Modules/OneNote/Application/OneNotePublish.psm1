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
        [string] $ID, # [in]BSTR bstrHierarchyID
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
    elseif ($PublishFormat -eq 'html') {
        $PublishFormat = 'pfHTML'
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
            $OneNotePage.Publish($ID, $Path, $PublishFormat, "")
        }
        else {
            Write-Host "Skipping Page: " $Path -ForegroundColor Yellow
        }
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($OneNotePage) | Out-Null
        Remove-Variable OneNotePage
    }
    catch {
        Write-Host $global:error -ForegroundColor Red
        Exit
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
        Write-Host $global:error -ForegroundColor Red
        Exit
    }
}

function Get-OneNoteEnrichPageCollection {
    <#
        Adds some extra properties to a page in page collection useful before a publish
        to make a hierarchy of nested pages

        On the way the page structure itself (3 level) should be structured are different opinions, so
        probably will be a paramter

        Also the possibility exists to auto unfold page structures (otherwise they will not be exported)

    #>
    param(
        [System.Array]$PageCollection
    )
    $count = $PageCollection.Count
    $Dir = "empty"
    $SubDir = "empty"
    $pageArray = New-Object -TypeName "System.Collections.ArrayList"
    $namesArray = New-Object -TypeName "System.Collections.ArrayList"
    try {
        for ($i = 0; $i -lt $count; $i++) {
            $page = $PageCollection[$i]

            # 1 copy over basic properties
            $pageObject = New-Object -TypeName PSObject
            Add-Member -InputObject $pageObject -MemberType NoteProperty -Name "Id" -Value ($page.ID) -Force
            Add-Member -InputObject $pageObject -MemberType NoteProperty -Name "Name" -Value (Remove-InvalidFileChars -Name $page.name) -Force
            Add-Member -InputObject $pageObject -MemberType NoteProperty -Name "DateTime" -Value ($page.dateTime) -Force
            Add-Member -InputObject $pageObject -MemberType NoteProperty -Name "LastModifiedTime" -Value ($page.lastModifiedTime) -Force
            Add-Member -InputObject $pageObject -MemberType NoteProperty -Name "PageLevel" -Value ($page.pageLevel) -Force

            # 2 Add property to indicate if a page has children
            $hasChildren = Get-OneNotePageHasChildren -PageCollection $pageCollection -ID ($pageObject.Id)
            if ($hasChildren) {
                Add-Member -InputObject $pageObject -MemberType NoteProperty -Name "HasChildren" -Value $true -Force
            }
            else {
                Add-Member -InputObject $pageObject -MemberType NoteProperty -Name "HasChildren" -Value $false -Force
            }

            # 3 Add path to indicate the path
            $path = ""
            if ($pageObject.PageLevel -eq 3) {
                $path = Join-Path -Path $Dir -ChildPath $SubDir
            }
            elseif ($pageObject.PageLevel -eq 2) {
                if ($pageObject.HasChildren) {
                    $SubDir = $pageObject.Name
                    $path = Join-Path -Path $Dir -ChildPath $SubDir
                }
                else {
                    $path = $Dir
                }
            }
            elseif ($pageObject.PageLevel -eq 1) {
                if ($pageObject.HasChildren) {
                    $Dir = $pageObject.Name
                    $path = $Dir
                }
                else {
                    $path = ""
                }
            }
            Add-Member -InputObject $pageObject -MemberType NoteProperty -Name "Path" -Value $path -Force

            # If the collection contains duplicate names, add and extension to one name
            $fullName = ""
            if ($null -ne $pageObject.Path -and "" -ne $pageObject.Path) {
                $fullName = Join-Path -Path ($pageObject.Path) -ChildPath ($pageObject.Name)
            }
            else {
                $fullName = $pageObject.Name
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
            Add-Member -InputObject $pageObject -MemberType NoteProperty -Name "FullName" -Value $fullName -Force

            # replace the enriched page
            $pageArray.Add($pageObject) | Out-Null
        }
        return $pageArray
    }
    catch {
        Write-Host "ERROR:" -ForegroundColor Red
        $global:error
        Exit
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
    $newName = $Name.Split([IO.Path]::GetInvalidFileNameChars()) -join '-'
    return (((($newName -replace "\s", "-") -replace "\[", "(") -replace "\]", ")").Substring(0, $(@{$true = 130; $false = $newName.length}[$newName.length -gt 150])))
}