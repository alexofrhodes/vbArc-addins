Attribute VB_Name = "Module1"


Sub MergeFileText(folderPath As String, Optional criteria As String = "*.txt")
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    Dim s As String
    Dim out As New Collection
    FilesAndOrFoldersInFolderOrZip folderPath, False, True, True, out, criteria
    For Each Item In out
        s = s & vbNewLine & ReadTXT(folderPath & Item)
    Next
    OverwriteTxt folderPath & "CombinedFileText.txt", s
End Sub

Function ProceduresOfTXT(FilePath As String, Optional nameOnly As Boolean) As String
    Dim var
    var = Split(ReadTXT(FilePath), Chr(10))
    Dim out
    out = JoinArrays(Filter(var, "Sub "), Filter(var, "Function "))
    If nameOnly = True Then
        Dim i As Long
        For i = LBound(out) To UBound(out)
            out(i) = Left(out(i), InStr(1, out(i), "(") - 1)
            out(i) = Replace(out(i), "Private ", "")
            out(i) = Replace(out(i), "Public ", "")
            out(i) = Replace(out(i), "Sub", "")
            out(i) = Replace(out(i), "Function ", "")
        Next
    End If
    SortArray out, LBound(out), UBound(out)
    Rem converted to string so I can use in immediate window:    ?ProceduresOfTXT(<PATH>)
    ProceduresOfTXT = Join(out, Chr(10))
End Function

Rem unsorted


Rem Folders
Public Function SelectFolder(Optional initFolder As String) As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = False
        .Title = "Select a folder"
        If FolderExists(initFolder) Then .InitialFileName = initFolder
        .Show
        If .SelectedItems.Count > 0 Then
            SelectFolder = .SelectedItems.Item(1)
        Else
        End If
    End With
End Function

Sub FoldersCreate(folderPath As String)
    Dim individualFolders() As String
    Dim tempFolderPath As String
    Dim arrayElement As Variant
    individualFolders = Split(folderPath, "\")
    For Each arrayElement In individualFolders
        tempFolderPath = tempFolderPath & arrayElement & "\"
        If FolderExists(tempFolderPath) = False Then
            MkDir tempFolderPath
        End If
    Next arrayElement
End Sub

Rem Files


Function GetFilePartName(fileNameWithExtension As String, Optional IncludeExtension As Boolean) As String
    If InStr(1, fileNameWithExtension, "\") > 0 Then
        GetFilePartName = Right(fileNameWithExtension, Len(fileNameWithExtension) - InStrRev(fileNameWithExtension, "\"))
    Else
        GetFilePartName = fileNameWithExtension
    End If
    If IncludeExtension = False Then GetFilePartName = Left(GetFilePartName, InStr(1, GetFilePartName, ".") - 1)
End Function

Function GetFilePartPath(fileNameWithExtension, Optional IncludeSlash As Boolean) As String
    GetFilePartPath = Left(fileNameWithExtension, InStrRev(fileNameWithExtension, "\") - 1 - IncludeSlash)
End Function

Public Function FFileDialog(Optional ByRef lDialogType As MsoFileDialogType = msoFileDialogFilePicker, _
                            Optional sTitle As String = "", _
                            Optional sInitFileName = "", _
                            Optional bMultiSelect As Boolean = False, _
                            Optional sFilter As String = "All Files,*.*") As String
    Dim out As String
    On Error GoTo Error_Handler
    Dim oFd                   As Object
    Dim vItems                As Variant
    Dim vFilter               As Variant
    Const msoFileDialogViewDetails = 2
    Set oFd = Application.FileDialog(lDialogType)
    With oFd
        If sTitle = "" Then
            Select Case lDialogType
                Case msoFileDialogFilePicker
                    .Title = "Browse for File"
                Case msoFileDialogFolderPicker
                    .Title = "Browse for Folder"
            End Select
        Else
            .Title = sTitle
        End If
        If sInitFileName <> "" Then .InitialFileName = sInitFileName
        .AllowMultiSelect = bMultiSelect
        .InitialView = msoFileDialogViewDetails
        If lDialogType <> msoFileDialogFolderPicker Then
            Call .Filters.Clear
            For Each vFilter In Split(sFilter, "~")
                Call .Filters.Add(Split(vFilter, ",")(0), Split(vFilter, ",")(1))
            Next vFilter
        End If
        If .Show = True Then
            For Each vItems In .SelectedItems
                If out = "" Then
                    out = vItems
                Else
                    out = out & "," & vItems
                End If
            Next
        End If
    End With
    FFileDialog = out
Error_Handler_Exit:
    On Error Resume Next
    If Not oFd Is Nothing Then Set oFd = Nothing
    Exit Function
Error_Handler:
    MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Error Source: fFileDialog" & vbCrLf & _
           "Error Description: " & Err.Description & _
           Switch(Erl = 0, "", Erl <> 0, vbCrLf & "Line No: " & Erl) _
           , vbOKOnly + vbCritical, "An Error has Occurred!"
    Resume Error_Handler_Exit
End Function

Public Function GetFilePath(Optional FileType As Variant, Optional multiSelect As Boolean) As Variant
    Dim blArray As Boolean
    Dim i As Long
    Dim strErrMsg As String, strTitle As String
    Dim varItem As Variant
    If Not IsMissing(FileType) Then
        blArray = IsArray(FileType)
        If Not blArray Then strErrMsg = "Please pass an array in the first parameter of this function!"
        If IsArrayAllocated(FileType) = False Then blArray = False
    End If
    If strErrMsg = vbNullString Then
        If multiSelect Then strTitle = "Choose one or more files" Else strTitle = "Choose file"
        With Application.FileDialog(msoFileDialogFilePicker)
            .InitialFileName = Environ("USERprofile") & "\Desktop\"
            .AllowMultiSelect = multiSelect
            .Filters.Clear
            If blArray Then .Filters.Add "File type", Replace("*." & Join(FileType, ", *."), "..", ".")
            .Title = strTitle
            If .Show <> 0 Then
                ReDim arrResults(1 To .SelectedItems.Count) As Variant
                For Each varItem In .SelectedItems
                    i = i + 1
                    arrResults(i) = varItem
                Next varItem
                GetFilePath = arrResults
            End If
        End With
    Else
        MsgBox strErrMsg, vbCritical, "Error!"
    End If
End Function

Rem TXT
Function TXTtoArray(sFile$)
    Rem https://newbedev.com/vb-vba-import-csv-to-array-code-example
    Rem VBA function to open a CSV file in memory and parse it to a 2D array without ever touching a worksheet:
    Dim c&, i&, j&, p&, D$, s$, rows&, cols&, a, R, v
    Const Q = """", QQ = Q & Q
    Const ENQ = ""
    Const ESC = ""
    Const COM = ","
    D = OpenTextFile$(sFile)
    If LenB(D) Then
        R = Split(Trim(D), vbCrLf)
        rows = UBound(R) + 1
        cols = UBound(Split(R(0), ",")) + 1
        ReDim v(1 To rows, 1 To cols)
        For i = 1 To rows
            s = R(i - 1)
            If LenB(s) Then
                If InStrB(s, QQ) Then s = Replace(s, QQ, ENQ)
                For p = 1 To Len(s)
                    Select Case Mid(s, p, 1)
                        Case Q:   c = c + 1
                        Case COM: If c Mod 2 Then Mid(s, p, 1) = ESC
                    End Select
                Next
                If InStrB(s, Q) Then s = Replace(s, Q, "")
                a = Split(s, COM)
                For j = 1 To cols
                    s = a(j - 1)
                    If InStrB(s, ESC) Then s = Replace(s, ESC, COM)
                    If InStrB(s, ENQ) Then s = Replace(s, ENQ, Q)
                    v(i, j) = s
                Next
            End If
        Next
        TXTtoArray = v
    End If
End Function

Rem insert string to txt file (not append, but on top)
Sub TxtPretend(FilePath As String, txt As String)
    Dim s As String
    s = ReadTXT(FilePath)
    OverwriteTxt FilePath, txt & Chr(10) & s
End Sub

Function TxtAppend(sFile As String, sText As String)
    On Error GoTo Err_Handler
    Dim iFileNumber           As Integer
    iFileNumber = FreeFile
    Open sFile For Append As #iFileNumber
    Print #iFileNumber, sText
    Close #iFileNumber
Exit_Err_Handler:
    Exit Function
Err_Handler:
    MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Error Source: Txt_Append" & vbCrLf & _
           "Error Description: " & Err.Description & _
           Switch(Erl = 0, "", Erl <> 0, vbCrLf & "Line No: " & Erl) _
           , vbOKOnly + vbCritical, "An Error has Occurred!"
    GoTo Exit_Err_Handler
End Function

Function OpenTextFile$(f)
    With CreateObject("ADODB.Stream")
        .Charset = "utf-8"
        .Open
        .LoadFromFile f
        OpenTextFile = .ReadText
        .Close
    End With
End Function

Function OverwriteTxt(sFile As String, sText As String)
    On Error GoTo Err_Handler
    Dim FileNumber As Integer
    FileNumber = FreeFile
    Open sFile For Output As #FileNumber
    Print #FileNumber, sText
    Close #FileNumber
Exit_Err_Handler:
    Exit Function
Err_Handler:
    MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Error Source: OverwriteTxt" & vbCrLf & _
           "Error Description: " & Err.Description, vbCritical, "An Error has Occurred!"
    GoTo Exit_Err_Handler
End Function

Function ReadTXT(sPath As String) As String
    Dim sTXT As String
    If Dir(sPath) = "" Then
        MsgBox "File was not found."
        Exit Function
    End If
    Open sPath For Input As #1
    Do Until EOF(1)
        Line Input #1, sTXT
        ReadTXT = ReadTXT & sTXT & vbLf
    Loop
    Close
    If Len(ReadTXT) = 0 Then
        ReadTXT = ""
    Else
        ReadTXT = Left(ReadTXT, Len(ReadTXT) - 1)
    End If
End Function


Public Function LoopAllFilesAndFolders(folderPath As String)
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    Dim objFSO As Scripting.FileSystemObject
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Dim objFile As Scripting.File
    Dim objFolder As Scripting.Folder
    Set objTopFolder = objFSO.GetFolder(folderPath)
    For Each objFile In objFolder.Files
    Next
    Dim objSubFolder As Scripting.Folder
    For Each objSubFolder In objFolder.SubFolders
        LoopAllFilesAndFolders objSubFolder.path
    Next
End Function

Sub FilesAndOrFoldersInFolderOrZipDemo()
    Dim out As New Collection
    FilesAndOrFoldersInFolderOrZip _
                                        FolderOrZipFilePath:="C:\Users\acer\Dropbox\SOFTWARE\EXCEL\00 Review", _
                                        LogFolders:=True, _
                                        LogFiles:=False, _
                                        ScanInSubfolders:=False, _
                                        out:=out
    dp out
End Sub
Function FilesAndOrFoldersInFolderOrZip(ByVal FolderOrZipFilePath As String, LogFolders As Boolean, LogFiles As Boolean, ScanInSubfolders As Boolean, out As Collection, Optional Filter As String = "*")
Dim oSh As New Shell
    Dim oFi As Object
    For Each oFi In oSh.Namespace(FolderOrZipFilePath).items
        If oFi.IsFolder Then
            If LogFolders Then
                out.Add oFi.path & "\"
            End If
            If ScanInSubfolders Then FilesAndOrFoldersInFolderOrZip oFi.path, LogFolders, LogFiles, ScanInSubfolders, out, Filter
        Else
            If LogFiles Then
                If UCase(oFi.Name) Like UCase(Filter) Then
                    out.Add oFi.path
                End If
            End If
        End If
    Next
    Set FilesAndOrFoldersInFolderOrZip = out
    Set oSh = Nothing
End Function
Public Sub UnzipToOwnFolder(ZippedFile As String, DeleteExistingFiles As Boolean, DeleteZip As Boolean)
 Rem for each cell in selection.cells: UnzipToOwnFolder cell.text,False,false :next
 

   Dim FileCollection As New Collection
    FilesAndOrFoldersInFolderOrZip ZippedFile, False, True, False, FileCollection
    Dim FolderCollection As New Collection
    FilesAndOrFoldersInFolderOrZip ZippedFile, True, False, False, FolderCollection
   
   Dim shell_app           As Object:     Set shell_app = CreateObject("Shell.Application")
Rem   Dim FilesInZip          As Long:        FilesInZip = shell_app.Namespace(CVar(ZippedFile)).items.Count
    Dim LastSlash            As Long:       LastSlash = InStrRev(ZippedFile, "\")
    Dim Dot                      As Long:      Dot = InStrRev(ZippedFile, ".")
    Dim ParentFolder       As String:     ParentFolder = Left(ZippedFile, LastSlash)
    Dim UnzipToFolder   As String
    If FolderCollection.Count = 1 And FileCollection.Count = 0 Then
        UnzipToFolder = ParentFolder
    ElseIf FolderCollection.Count > 1 Or FileCollection.Count > 0 Then
        UnzipToFolder = Left(ZippedFile, Dot - 1) & "\"
        If DeleteExistingFiles Then
            If FolderExists(UnzipToFolder) Then RecycleSafe UnzipToFolder
        End If
        FoldersCreate UnzipToFolder
    End If
    shell_app.Namespace(CVar(UnzipToFolder)).copyhere shell_app.Namespace(CVar(ZippedFile)).items
    If DeleteZip Then RecycleSafe ZippedFile
    Set shell_app = Nothing
End Sub

Public Sub UnzipAllInFolder(source_folder As String)
    Dim current_zip_file As String
    current_zip_file = Dir(source_folder & "\*.zip")
    If Len(current_zip_file) = 0 Then
        MsgBox "No zip files found!", vbExclamation
        Exit Sub
    End If
    Dim zip_folder As String
    zip_folder = source_folder & "\unzipped"
    Dim error_message As String
    If Not create_temp_zip_folder(zip_folder, error_message) Then
        MsgBox error_message, vbCritical, "Error"
        Exit Sub
    End If
    Dim shell_app As Object
    Set shell_app = CreateObject("Shell.Application")
    Do While Len(current_zip_file) > 0
        shell_app.Namespace(CVar(zip_folder)).copyhere shell_app.Namespace(source_folder & "\" & current_zip_file).items
        current_zip_file = Dir
    Loop
    Set shell_app = Nothing
End Sub

Function create_temp_zip_folder(ByVal zip_folder As String, ByRef error_message As String) As Boolean
    On Error GoTo Error_Handler
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(zip_folder) Then
        fso.DeleteFolder zip_folder, True
    End If
    fso.CreateFolder zip_folder
    create_temp_zip_folder = True
    Set fso = Nothing
    Exit Function
Error_Handler:
    error_message = "Error " & Err.Number & ":" & vbCrLf & vbCrLf & Err.Description
End Function

Sub SelectControItemsByFilter(lbox As MSForms.ListBox, criteria As String)
    DeselectAll lbox
    If criteria = "" Then Exit Sub
    For i = 0 To lbox.ListCount - 1
        If UCase(lbox.List(i, 1)) Like "*" & UCase(criteria) & "*" Then
            lbox.Selected(i) = True
        End If
    Next i
End Sub
Sub DeselectAll(lbox As MSForms.ListBox)
    If lbox.ListCount = 0 Then Exit Sub
    For i = 0 To lbox.ListCount - 1
        lbox.Selected(i) = False
    Next i
End Sub
Function ListboxSelectedValues(listboxCollection As Variant, CollectionToArray As Boolean) As Variant
    Dim i As Long
Dim listItem As Long
Dim selectedCollection As New Collection
Dim listboxCount As Long
'if arguement passed is collection of listboxes
If TypeName(listboxCollection) = "Collection" Then
    For listboxCount = 1 To listboxCollection.Count
        If listboxCollection(listboxCount).ListCount > 0 Then
            For listItem = 0 To listboxCollection(listboxCount).ListCount - 1
                If listboxCollection(listboxCount).Selected(listItem) Then
                    selectedCollection.Add CStr(listboxCollection(listboxCount).List(listItem))
                End If
            Next listItem
        End If
    Next listboxCount
'if arguement passed is single Listbox
Else
        If listboxCollection.ListCount > 0 Then
        For i = 0 To listboxCollection.ListCount - 1
            If listboxCollection.Selected(i) Then
                selectedCollection.Add CStr(listboxCollection(listboxCount).List(listItem))
            End If
        Next i
    End If
End If
If CollectionToArray = True Then
    ListboxSelectedValues = CollectionToArray(selectedCollection)
Else
    Set ListboxSelectedValues = selectedCollection
End If
End Function
Function ListboxSelectedCount(listboxCollection As Variant) As Long
    Dim i As Long
Dim listItem As Long
Dim selectedCollection As New Collection
Dim listboxCount As Long
'if arguement passed is collection of listboxes
If TypeName(listboxCollection) = "Collection" Then
    For listboxCount = 1 To listboxCollection.Count
        If listboxCollection(listboxCount).ListCount > 0 Then
            For listItem = 0 To listboxCollection(listboxCount).ListCount - 1
                If listboxCollection(listboxCount).Selected(listItem) Then
                    selectedCount = selectedCount + 1
                End If
            Next listItem
        End If
    Next listboxCount
'if arguement passed is single Listbox
Else
        If listboxCollection.ListCount > 0 Then
        For i = 0 To listboxCollection.ListCount - 1
            If listboxCollection.Selected(i) Then
                selectedCount = selectedCount + 1
            End If
        Next i
    End If
End If
ListboxSelectedCount = selectedCount
End Function
Function ListboxSelectedIndexes(lbox As MSForms.ListBox, TransformCollectionToArray As Boolean) As Variant
'listboxes start at 0
Dim i As Long
Dim selectedIndexes As New Collection
    If lbox.ListCount > 0 Then
        For i = 0 To lbox.ListCount - 1
            If lbox.Selected(i) Then selectedIndexes.Add i
        Next i
    End If
If TransformCollectionToArray = True Then
    ListboxSelectedIndexes = CollectionToArray(selectedIndexes)
Else
    Set ListboxSelectedIndexes = selectedIndexes
End If
End Function
Function CollectionToArray(c As Collection) As Variant
    Dim a() As Variant: ReDim a(0 To c.Count - 1)
    Dim i As Long
    For i = 1 To c.Count
        a(i - 1) = c.Item(i)
    Next
    CollectionToArray = a
End Function
Function isFDU(path) As String
    Dim retval
    retval = "I"
    If (retval = "I") And FileExists(path) Then retval = "F"
    If (retval = "I") And FolderExists(path) Then retval = "D"
    If (retval = "I") And HttpExists(path) Then retval = "U"
    ' I => Invalid | F => File | D => Directory | U => Valid Url
    isFDU = retval
End Function
Function FileExists(ByVal strFile As String, Optional bFindFolders As Boolean) As Boolean
    'Purpose:   Return True if the file exists, even if it is hidden.
    'Arguments: strFile: File name to look for. Current directory searched if no path included.
    '           bFindFolders. If strFile is a folder, FileExists() returns False unless this argument is True.
    'Note:      Does not look inside subdirectories for the file.
    'Author:    Allen Browne. http://allenbrowne.com June, 2006.
    Dim lngAttributes As Long

    'Include read-only files, hidden files, system files.
    lngAttributes = (vbReadOnly Or vbHidden Or vbSystem)
    If bFindFolders Then
        lngAttributes = (lngAttributes Or vbDirectory) 'Include folders as well.
    Else
        'Strip any trailing slash, so Dir does not look inside the folder.
        Do While Right$(strFile, 1) = "\"
            strFile = Left$(strFile, Len(strFile) - 1)
        Loop
    End If
    'If Dir() returns something, the file exists.
    On Error Resume Next
    FileExists = (Len(Dir(strFile, lngAttributes)) > 0)
End Function
Function FolderExists(ByVal strPath As String) As Boolean
    On Error Resume Next
    FolderExists = ((GetAttr(strPath) And vbDirectory) = vbDirectory)
End Function
Function TrailingSlash(varIn As Variant) As String
    If Len(varIn) > 0 Then
        If Right(varIn, 1) = "\" Then
            TrailingSlash = varIn
        Else
            TrailingSlash = varIn & "\"
        End If
    End If
End Function
Function HttpExists(ByVal sURL As String) As Boolean
    Dim oXHTTP As Object
    Set oXHTTP = CreateObject("MSXML2.XMLHTTP")
    If Not UCase(sURL) Like "HTTP:*" Then
    sURL = "http://" & sURL
    End If
    On Error GoTo haveError
    oXHTTP.Open "HEAD", sURL, False
    oXHTTP.send
    HttpExists = IIf(oXHTTP.Status = 200, True, False)
    Exit Function
haveError:
    Debug.Print Err.Description
    HttpExists = False
End Function

Function whichOption(frame As Variant, controlType As String) As Variant
Dim out As New Collection
Dim Control As MSForms.Control
    For Each Control In frame.Controls
        If UCase(TypeName(Control)) = UCase(controlType) Then
            If Control.Value = True Then
                out.Add Control
                If TypeName(frame) = "Frame" Then Exit For
            End If
        End If
    Next
    
    If out.Count = 1 Then
        Set whichOption = out(1)
    ElseIf out.Count > 1 Then
        Set whichOption = out
    End If
End Function
