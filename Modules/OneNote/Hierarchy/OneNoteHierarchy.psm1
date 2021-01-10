function Get-OneNoteHierarchy {
    <#
        .SYNOPSIS
            Gets the OneNote Hierarchy
            https://docs.microsoft.com/en-us/office/client-developer/onenote/application-interface-onenote
            see: https://docs.microsoft.com/en-us/office/dev/add-ins/images/onenote-om.png for structure
            Gets the notebook node hierarchy structure, starting from the node you specify (all notebooks or a single notebook,
            section group, or section), and extending downward to all descendants at the level you specify.
    #>
    param(
        [string]$bstrStartNodID = ""
        <#
            [in]BSTR bstrPath
            The path that you want to open. For a notebook, or for a section group in a notebook, bstrPath can be a folder path
            or the path to an .one section file. If you specify the path to an .one section file, you must include the .one extension
            on the file-path string.

            If we do not specify a path it takes the open OneNote Application

            [in]BSTR bstrRelativeToObjectID
            The OneNote ID of the parent object (notebook or section group) under which you want the new object to open. If the bstrPath
            parameter is an absolute path, you can pass an empty string ("") for bstrRelativeToObjectID. Alternatively, you
            can pass the object ID of the notebook or section group that should contain the object (section or section group)
            that you want to create, and then specify the file name (for example, section1.one) of the object that you want to
            create under that parent object.

            pbstrObjectID – (Output parameter)
            The object ID that OneNote returns for the notebook, section group, or section that the OpenHierarchy method opens.
            This parameter is a pointer to the string into which you want the method to write the ID.

            cftlfNotExist – (Optional) An enumerated value from the CreateFileType enumeration. If you pass a value for
            cftIfNotExist, the OpenHierarchy method creates the section group or section file at the specified path only
            if the file does not already exist.
            (not implemented)
        #>
    )
    try {
        [xml]$ObjectId = $null
        $OneNote = New-Object -ComObject OneNote.Application
        $OneNote.GetHierarchy("", [Microsoft.Office.InterOp.OneNote.HierarchyScope]::hsPages, [ref]$ObjectId)
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($OneNote)
        Remove-Variable OneNote
        return $ObjectId
    }
    catch
    {
        Write-Host $global:error -ForegroundColor Red
        Exit
    }
}