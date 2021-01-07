#
# import all modules
#
(Get-ChildItem -Path "$PSScriptRoot\Modules" -Recurse -Filter '*.psm1' -Verbose).FullName | ForEach-Object { Import-Module $_ -Force }

function RenameImages {
  <#
    Rename images to have unique names  - NoteName-Image#-HHmmssff.xyz
  #>
  param (
    [string]$MdPath,
    [string]$MediaPath,
    [string]$PageName,
    [string]$LevelsPrefix
  )
  try {
    $timeStamp = (Get-Date -Format HHmmssff).ToString()
    $timeStamp = $timeStamp.replace(':', '')
    $images = Get-ChildItem -Path "$($MediaPath)/media" -Include "*.png", "*.gif", "*.jpg", "*.jpeg" -Recurse -ErrorAction SilentlyContinue | Where-Object { $_.Name.SubString(0, 5) -match "image" }
    foreach ($image in $images) {
      $newimageName = "$($PageName.SubString(0,[math]::min(30,$PageName.length)))-$($image.BaseName)-$($timeStamp)$($image.Extension)"
      Rename-Item -Path "$($image.FullName)" -NewName $newimageName -ErrorAction SilentlyContinue
      ((Get-Content -path $MdPath  -Raw).Replace("$($image.Name)", "$($newimageName)")) | Set-Content -Path $MdPath
    }
    # Change MD file Image Path References in Markdown
    ((Get-Content -path $MdPath  -Raw).Replace("$($MediaPath.Replace("\","\\"))", "$($LevelsPrefix)")) | Set-Content -Path $MdPath
    # Change MD file Image Path References in HTML
    ((Get-Content -path $MdPath  -Raw).Replace("$($MediaPath)", "$($LevelsPrefix)")) | Set-Content -Path $MdPath
  }
  catch {
    Write-Host $Error -ForegroundColor Red
    Exit
  }
}

function ClearSpacesInPage {
  <#
    Clear double spaces from bullets and nonbreaking spaces from blank lines
  #>
  Param (
    [string]$MdPath,
    [bool]$ClearSpaces

  )
  if ($ClearSpaces -eq $false) {
    try {
      ((Get-Content -path $MdPath -Raw -encoding utf8).Replace(">", "").Replace("<", "").Replace([char]0x00A0, [char]0x000A).Replace([char]0x000A, [char]0x000A).Replace("`r`n`r`n", "`r`n")) | Set-Content -Path $MdPath
    }
    catch {
      Write-Host "Error while clearing double spaces from file" $MdPath " : $($Error[0].ToString())" -ForegroundColor Red
    }
  }
}

function AddYamlToPage {
  Param(
    [string]$MdPath,
    [System.Xml.XmlElement]$Page
  )
  $orig = Get-Content -path $MdPath
  $orig[0] = "# $($Page.name)"
  $insert1 = "$($Page.dateTime)"
  $insert1 = [Datetime]::ParseExact($insert1, 'yyyy-MM-ddTHH:mm:ss.fffZ', $null)
  $insert1 = $insert1.ToString("yyyy-MM-dd HH:mm:ss
        ")
  $insert2 = "---"
  Set-Content -Path $MdPath -Value $orig[0..0], $insert1, $insert2, $orig[6..$orig.Length]
}

function ProcessPageFileObjects {
  Param(
    [string]$PageId,
    [System.Xml.XmlElement]$Page,
    [string]$MediaPath,
    [string]$MdPath
  )
  $files = Get-OneNotePageInsertedFileObjects -PageId $pageId -Page $Page
  foreach ($file in $files) {
    $destfilename = ""
    $path = "$($MediaPath)\media"
    New-Dir -Path $path
    try {
      $destfilename = $file.InsertedFile.preferredName | Remove-InvalidFileNameCharsInsertedFiles
      Copy-Item -Path "$($file.InsertedFile.pathCache)" -Destination "$($path)\$($destfilename)" -Force
    }
    catch {
      Write-Host "Error while copying file object '$($file.InsertedFile.preferredName)' for page '$($Page.name)': $($Error[0].ToString())" -ForegroundColor Red
    }
    # Change MD file Object Name References
    try {
      $pageinsertedfile2 = $destfilename.Replace("$", "\$").Replace("^", "\^").Replace("'", "\'")
      ReplaceStringInFile -File $MdPath -StringToBeReplaced "$($pageinsertedfile2)" -StringThatWillReplaceIt "[$($destfilename)]($($path)/$($destfilename))"
    }
    catch {
      Write-Host "Error while renaming file object name references to '$($file.InsertedFile.preferredName)' for file '$($page.name)': $($Error[0].ToString())" -ForegroundColor Red
    }
  }
}

function Export-OneNotePageCollection {
  <#
    Process pagecollections in sections
  #>
  param(
    [Object]$Config,
    [System.Array]$PageCollection,
    [String]$Path,
    [Int]$Level
  )
  try {
    $enrichedPageCollection = Get-OneNoteEnrichPageCollection -PageCollection $pageCollection
    foreach ($page in $enrichedPageCollection) {
      Write-Host " --------------- PAGE ------------------------ "
      $page
      $pagename = Get-OneNotePageName -Page $page

      $fileNameWithoutExtension = (Join-Path $Config.ExportRootPath -ChildPath "docx" |
        Join-Path -ChildPath $Path | Join-Path -ChildPath $page.Path | Join-Path  -ChildPath $pagename)

      $wordDocumentPath = $fileNameWithoutExtension + ".docx"
      if ([System.IO.File]::Exists($wordDocumentPath)) {
        $wordDocumentPath = $fileNameWithoutExtension + $page.ID + ".docx"
      }

      $mdDocumentPath = $fileNameWithoutExtension + ".md"
      if ([System.IO.File]::Exists($mdDocumentPath)) {
        $mdDocumentPath = $fileNameWithoutExtension + $page.ID + ".md"
      }

      Write-Host "word: " $wordDocumentPath
      Write-Host "md: " $mdDocumentPath

      <#
      $relativeRoot = Join-Path -Path $Config.ExportRootPath -ChildPath $Path
      $pageRoot = Join-Path  -Path pageRoot -ChildPath $page.Path
      $fullPathWithoutExtension = Join-Path  -Path $pageRoot -ChildPath $pagename
      $docNam =

      $fullexportdirpath = $Path
      $fullfilepathwithoutextension = "$($fullexportdirpath)\$($pagename)"

      $levelsprefix = "../" * ($levelsfromroot) + ".."

      # set media location (central media folder at notebook-level or adjacent to .md file) based on initial user prompt
      if ($Config.CentralMediaFolder -eq $true) {
        $mediaPath = $fullexportdirpath
        $levelsprefix = ""
      }
      else {
        $mediaPath = $NotebookFilePath
      }
      #>

      #Invoke-OneNotePagePublish -ID ($page.ID) -Path $wordDocumentPath -Overwrite $true

      <#
      Invoke-ConvertDocxToMd -OutputFormat $converter -Inputfile $wordDocumentPath -OutputFile "$fullfilepathwithoutextension.md" -MediaPath $mediaPath
      ProcessPageFileObjects -PageId $page.ID -Page $_ -MediaPath $MediaPath -MdPath $mdDocumentPath
      AddYamlToPage -MdPath $mdDocumentPath -Page $_
      ClearSpacesInPage -MdPath $mdDocumentPath -ClearSpaces $true # todo replace with variable
      RenameImages -MdPath $mdDocumentPath -MediaPath $MediaPath -PageName $pagename -LevelsPrefix $levelsprefix
      if ($config.ClearEscape) {
        ReplaceStringInFile -File $wordDocumentPath -StringToBeReplaced "\" -StringThatWillReplaceIt ""
      }
      if ($Config.KeepWordFiles -eq $false) {
        Remove-File -File $wordDocumentPath
      }
      #>
    }
  }
  catch {
    Write-Host $Error -ForegroundColor Red
    Exit
  }
}

function Export-OneNoteSectionCollection {
  <#
    Handle SectionCollections
  #>
  param(
    [Object]$Config,
    [System.Array] $SectionCollection,
    [String]$Path,
    [Int]$Level = 0
  )
  try {
    foreach ($section in $SectionCollection) {
      $sectionFileName = Get-OneNoteSectionName -Section $section
      $dir = Join-Path $Path -ChildPath $sectionFileName
      Write-Host "Exporting Section $dir" -ForegroundColor Green
      $pageCollection = Get-OneNoteSectionPageCollection -Section $section
      Export-OneNotePageCollection -PageCollection $pageCollection -Path $dir -Level $Level -Config $Config
    }
  }
  catch {
    Write-Host $Error -ForegroundColor Red
    Exit
  }
}

function Export-OneNoteSectionGroupSection {
  <#
    Exports One item in a SectionGroup Collection
  #>
  param (
    [Object]$Config,
    [System.Xml.XmlElement]$SectionGroup,
    [int]$Level,
    [string]$type
  )
  try {
    $sectionGroupName = Get-OneNoteSectionGroupName -SectionGroup $SectionGroup
    $dir = Join-Path $Path -ChildPath $sectionGroupName
    Write-Host "Exporting $type $dir" -ForegroundColor Green
    if ($type -eq "Section") {
      Export-OneNoteSectionCollection -Config $Config -SectionCollection $SectionGroup.Section -Path $dir -Level ($Level+1)
    } elseif ($type -eq "SectionGroup") {
      Export-OneNoteSectionGroupCollection -Config $Config -SectionGroupCollection $SectionGroup.SectionGroup -Path $dir -Level ($Level+1)
    }
  }
  catch {
    Write-Host $Error -ForegroundColor Red
    Exit
  }
}

function Export-OneNoteSectionGroupCollection {
  <#
        Export SectionGroup Collection. A section group can contain an entry "Section" or an entry "SectionGroup"
  #>
  param (
    [Object]$Config,
    [System.Array]$SectionGroupCollection,
    [String]$Path,
    [Int]$Level = 0
  )
  try {
    if ($SectionGroupCollection.Count -gt 0 -and $null -ne $SectionGroupcollection) {
      Write-Host "Exporting New SectionGroup $Path" -ForegroundColor Green
      $sectionGroupItemsWithSection = Get-OneNoteSectionGroupCollectionSectionCollection -SectionGroup $SectionGroupCollection
      Write-Host "Parsing Entries Section in SectionGroup" -ForegroundColor Green
      foreach ($sectionGroup in $sectionGroupItemsWithSection) {
        Export-OneNoteSectionGroupSection -Config $config -SectionGroup $sectionGroup -Path $Path -Level $Level -type "Section"
      }
      # get all entries that contain an entry "sectiongroup"
      Write-Host "Parsing Entries SectionGroup in SectionGroup" -ForegroundColor Green
      $sectionGroupItemsWithSectionGroup = Get-OneNoteSectionGroupCollectionSectionGroupCollection -SectionGroup $sectionGroupCollection
      foreach ($sectionGroup in $sectionGroupItemsWithSectionGroup ) {
        Export-OneNoteSectionGroupSection -Config $config -SectionGroup $sectionGroup -Path $Path -Level $Level -type "SectionGroup"
      }
    }
  }
  catch {
    Write-Host $Error -ForegroundColor Red
    Exit
  }
}

function Export-OneNoteNoteBook {
  <#
    Export one OneNote Notebook
  #>
  param (
    [Object]$Config,
    [System.Xml.XmlElement]$Notebook
  )
  try {
    Write-Host "Exporting Sections and SectionGroups from Notebook $notebookFileName" -ForegroundColor Green
    $notebookFileName = Get-OneNoteNotebookCleanFileName -Notebook $Notebook
    $sectionCollection = Get-OneNoteNotebookSectionCollection -Notebook $Notebook
    $sectionGroupCollection = Get-OneNoteNotebookSectionGroupCollection -Notebook $Notebook -Include_Recyclebin $false
    if ($null -ne $sectionCollection) {
      Export-OneNoteSectionCollection -SectionCollection $sectionCollection -Path $notebookFileName -Config $Config
    }
    if ($null -ne $sectionGroupCollection) {
      Export-OneNoteSectionGroupCollection -SectionGroupCollection $sectionGroupCollection -Path $notebookFileName -Config $Config
    }
  }
  catch {
    Write-Host $Error -ForegroundColor Red
    Exit
  }
}

function Export-OneNoteNotebookCollection {
  <#
    Export Collection of Notebooks
  #>
  param (
    [Object]$Config,
    [System.Xml.XmlElement]$NotebookCollection
  )
  try {
    $count = Get-OneNoteNotebookCollectionCount -NotebookCollection $NotebookCollection
    for($item=0; $item -lt $count; $item++) {
      Write-Host "Exporting Notebook $($item+1) out of $count" -ForegroundColor Green
      $notebook = Get-OneNoteNotebook -NotebookCollection $NotebookCollection -NotebookItem $item
      Write-Host "Exporting Notebook:" -ForegroundColor Green
      $notebook
      Export-OneNoteNoteBook -Notebook $notebook -Config $Config
    }
  }
  catch {
    Write-Host $Error -ForegroundColor Red
    Exit
  }
}

function Export-OneNoteHierachy {
  <#
    Exports OneNote Hierarchy
  #>
  param (
    [Object]$Config
  )
  try {
    $hierarchy = Get-OneNoteHierarchy
    $notebookCollection = Get-OneNoteNoteBookCollection -Hierarchy $hierarchy
    Export-OneNoteNotebookCollection -NotebookCollection $notebookCollection -Config $config
  }
  catch {
    Write-Host $Error -ForegroundColor Red
    Exit
  }
}

Export-OneNoteHierachy -Config (Get-Config -path "$PSScriptRoot\Config\onenoteexport.cfg")
