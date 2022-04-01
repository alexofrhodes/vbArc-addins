Attribute VB_Name = "mRecycle"
Rem recycle
Private Declare PtrSafe Function SHFileOperation Lib "shell32.dll" Alias _
"SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Private Declare PtrSafe Function PathIsNetworkPath Lib "shlwapi.dll" _
Alias "PathIsNetworkPathA" ( _
ByVal pszPath As String) As Long
Private Declare PtrSafe Function GetSystemDirectory Lib "kernel32" _
Alias "GetSystemDirectoryA" ( _
ByVal lpBuffer As String, _
ByVal nSize As Long) As Long
Private Declare PtrSafe Function SHEmptyRecycleBin _
Lib "Shell32" Alias "SHEmptyRecycleBinA" _
(ByVal hWnd As Long, _
ByVal pszRootPath As String, _
ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function PathIsDirectory Lib "shlwapi" (ByVal pszPath As String) As Long
Private Const FO_DELETE = &H3
Private Const FOF_ALLOWUNDO = &H40
Private Const FOF_NOCONFIRMATION = &H10
Private Const MAX_PATH As Long = 260
Private Type SHFILEOPSTRUCT
    hWnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Boolean
    hNameMappings As Long
    lpszProgressTitle As String
End Type

Public Function RecycleFile(FileName As String) As Boolean
    Dim SHFileOp As SHFILEOPSTRUCT
    Dim Res As Long
    If Dir(FileName, vbNormal) = vbNullString Then
        RecycleFile = True
        Exit Function
    End If
    With SHFileOp
        .wFunc = FO_DELETE
        .pFrom = FileName
        .fFlags = FOF_ALLOWUNDO
        .fFlags = FOF_ALLOWUNDO + FOF_NOCONFIRMATION
    End With
    Res = SHFileOperation(SHFileOp)
    If Res = 0 Then
        RecycleFile = True
    Else
        RecycleFile = False
    End If
End Function

Public Function Recycle(FileSpec As String, Optional ErrText As String) As Boolean
    Dim SHFileOp As SHFILEOPSTRUCT
    Dim Res As Long
    Dim sFileSpec As String
    ErrText = vbNullString
    sFileSpec = FileSpec
    If InStr(1, FileSpec, ":", vbBinaryCompare) = 0 Then
        ErrText = "'" & FileSpec & "' is not a fully qualified name on the local machine"
        Recycle = False
        Exit Function
    End If
    If Dir(FileSpec, vbDirectory) = vbNullString Then
        ErrText = "'" & FileSpec & "' does not exist"
        Recycle = False
        Exit Function
    End If
    If Right(sFileSpec, 1) = "\" Then
        sFileSpec = Left(sFileSpec, Len(sFileSpec) - 1)
    End If
    With SHFileOp
        .wFunc = FO_DELETE
        .pFrom = sFileSpec
        .fFlags = FOF_ALLOWUNDO
        .fFlags = FOF_ALLOWUNDO + FOF_NOCONFIRMATION
    End With
    Res = SHFileOperation(SHFileOp)
    If Res = 0 Then
        Recycle = True
    Else
        Recycle = False
    End If
End Function

Public Function RecycleSafe(FileSpec As String, Optional ByRef ErrText As String) As Boolean
    Dim ThisWorkbookFullName As String
    Dim ThisWorkbookPath As String
    Dim WindowsFolder As String
    Dim SystemFolder As String
    Dim ProgramFiles As String
    Dim MyDocuments As String
    Dim Desktop As String
    Dim ApplicationPath As String
    Dim Pos As Long
    Dim ShellObj As Object
    Dim sFileSpec As String
    Dim SHFileOp As SHFILEOPSTRUCT
    Dim Res As Long
    Dim FileNum As Integer
    sFileSpec = FileSpec
    If InStr(1, FileSpec, ":", vbBinaryCompare) = 0 Then
        RecycleSafe = False
        ErrText = "'" & FileSpec & "' is not a fully qualified name on the local machine"
        Exit Function
    End If
    If Dir(FileSpec, vbDirectory) = vbNullString Then
        RecycleSafe = False
        ErrText = "'" & FileSpec & "' does not exist"
        Exit Function
    End If
    If Right(sFileSpec, 1) = "\" Then
        sFileSpec = Left(sFileSpec, Len(sFileSpec) - 1)
    End If
    ThisWorkbookFullName = ThisWorkbook.FullName
    ThisWorkbookPath = ThisWorkbook.path
    SystemFolder = String$(MAX_PATH, vbNullChar)
    GetSystemDirectory SystemFolder, Len(SystemFolder)
    SystemFolder = Left(SystemFolder, InStr(1, SystemFolder, vbNullChar, vbBinaryCompare) - 1)
    Pos = InStrRev(SystemFolder, "\")
    If Pos > 0 Then
        WindowsFolder = Left(SystemFolder, Pos - 1)
    End If
    Pos = InStr(1, Application.path, "\", vbBinaryCompare)
    Pos = InStr(Pos + 1, Application.path, "\", vbBinaryCompare)
    ProgramFiles = Left(Application.path, Pos - 1)
    ApplicationPath = Application.path
    On Error Resume Next
    Err.Clear
    Set ShellObj = CreateObject("WScript.Shell")
    If ShellObj Is Nothing Then
        RecycleSafe = False
        ErrText = "Error Creating WScript.Shell. " & CStr(Err.Number) & ": " & Err.Description
        Exit Function
    End If
    MyDocuments = ShellObj.SpecialFolders("MyDocuments")
    Desktop = ShellObj.SpecialFolders("Desktop")
    Set ShellObj = Nothing
    If (sFileSpec Like "?*:") Or (sFileSpec Like "?*:\") Then
        RecycleSafe = False
        ErrText = "File Specification is a root directory."
        Exit Function
    End If
    If (InStr(1, sFileSpec, "*", vbBinaryCompare) > 0) Or (InStr(1, sFileSpec, "?", vbBinaryCompare) > 0) Then
        RecycleSafe = False
        ErrText = "File specification contains wildcard characters"
        Exit Function
    End If
    If StrComp(sFileSpec, ThisWorkbookFullName, vbTextCompare) = 0 Then
        RecycleSafe = False
        ErrText = "File specification is the same as this workbook."
        Exit Function
    End If
    If StrComp(sFileSpec, ThisWorkbookPath, vbTextCompare) = 0 Then
        RecycleSafe = False
        ErrText = "File specification is this workbook's path"
        Exit Function
    End If
    If StrComp(ThisWorkbook.FullName, sFileSpec, vbTextCompare) = 0 Then
        RecycleSafe = False
        ErrText = "File specification is this workbook."
        Exit Function
    End If
    If StrComp(sFileSpec, SystemFolder, vbTextCompare) = 0 Then
        RecycleSafe = False
        ErrText = "File specification is the System Folder"
        Exit Function
    End If
    If StrComp(sFileSpec, WindowsFolder, vbTextCompare) = 0 Then
        RecycleSafe = False
        ErrText = "File specification is the Windows folder"
        Exit Function
    End If
    If StrComp(sFileSpec, Application.path, vbTextCompare) = 0 Then
        RecycleSafe = False
        ErrText = "File specification is Application Path"
        Exit Function
    End If
    If StrComp(sFileSpec, MyDocuments, vbTextCompare) = 0 Then
        RecycleSafe = False
        ErrText = "File specification is MyDocuments"
        Exit Function
    End If
    If StrComp(sFileSpec, Desktop, vbTextCompare) = 0 Then
        RecycleSafe = False
        ErrText = "File specification is Desktop"
        Exit Function
    End If
    If (GetAttr(sFileSpec) And vbSystem) <> 0 Then
        RecycleSafe = False
        ErrText = "File specification is a System entity"
        Exit Function
    End If
    If PathIsDirectory(sFileSpec) = 0 Then
        FileNum = FreeFile()
        On Error Resume Next
        Err.Clear
        Open sFileSpec For Input Lock Read As #FileNum
        If Err.Number <> 0 Then
            Close #FileNum
            RecycleSafe = False
            ErrText = "File in use: " & CStr(Err.Number) & "  " & Err.Description
            Exit Function
        End If
        Close #FileNum
    End If
    With SHFileOp
        .wFunc = FO_DELETE
        .pFrom = sFileSpec
        .fFlags = FOF_ALLOWUNDO
        .fFlags = FOF_ALLOWUNDO + FOF_NOCONFIRMATION
    End With
    Res = SHFileOperation(SHFileOp)
    If Res = 0 Then
        RecycleSafe = True
    Else
        RecycleSafe = False
    End If
End Function

Public Function EmptyRecycleBin(Optional DriveRoot As String = vbNullString) As Boolean
    Const SHERB_NOCONFIRMATION = &H1
    Const SHERB_NOPROGRESSUI = &H2
    Const SHERB_NOSOUND = &H4
    Dim Res As Long
    If DriveRoot <> vbNullString Then
        If PathIsNetworkPath(DriveRoot) <> 0 Then
            MsgBox "You can't empty the Recycle Bin of a network drive."
            Exit Function
        End If
    End If
    Res = SHEmptyRecycleBin(hWnd:=0&, _
                            pszRootPath:=DriveRoot, _
                            dwFlags:=SHERB_NOCONFIRMATION + _
                                      SHERB_NOPROGRESSUI + _
                                      SHERB_NOSOUND)
    If Res = 0 Then
        EmptyRecycleBin = True
    Else
        EmptyRecycleBin = False
    End If
End Function

Rem End of Recycle

