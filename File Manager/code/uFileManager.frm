VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uFileManager 
   Caption         =   "Drag and drop files or folders to the listbox"
   ClientHeight    =   5232
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6864
   OleObjectBlob   =   "uFileManager.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uFileManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type POINTAPI
  X As Long
  Y As Long
End Type

#If VBA7 Then
    Private Type MSG
        hWnd As LongPtr
        message As Long
        wParam As LongPtr
        lParam As LongPtr
        time As Long
        pt As POINTAPI
    End Type

    Private Declare PtrSafe Function GetMessage Lib "User32" Alias "GetMessageA" (lpMsg As MSG, ByVal hWnd As LongPtr, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
    Private Declare PtrSafe Function DispatchMessage Lib "User32" Alias "DispatchMessageA" (lpMsg As MSG) As LongPtr
    Private Declare PtrSafe Function TranslateMessage Lib "User32" (lpMsg As MSG) As Long
    Private Declare PtrSafe Function WindowFromAccessibleObject Lib "oleacc" (ByVal pacc As IAccessible, phwnd As LongPtr) As Long
    Private Declare PtrSafe Function IsWindow Lib "User32" (ByVal hWnd As LongPtr) As Long
    Private Declare PtrSafe Sub DragAcceptFiles Lib "shell32.dll" (ByVal hWnd As LongPtr, ByVal fAccept As Long)
    Private Declare PtrSafe Sub DragFinish Lib "shell32.dll" (ByVal HDROP As LongPtr)
    Private Declare PtrSafe Function DragQueryFile Lib "shell32.dll" Alias "DragQueryFileA" (ByVal HDROP As LongPtr, ByVal UINT As Long, ByVal lpStr As String, ByVal ch As Long) As Long
#Else

    Private Type MSG
        hWnd As Long
        message As Long
        wParam As Long
        lParam As Long
        time As Long
        pt As POINTAPI
    End Type

    Private Declare Function GetMessage Lib "user32" Alias "GetMessageA" (lpMsg As MSG, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
    Private Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As MSG) As Long
    Private Declare Function TranslateMessage Lib "user32" (lpMsg As MSG) As Long
    Private Declare Function WindowFromAccessibleObject Lib "oleacc" (ByVal pacc As IAccessible, phwnd As Long) As Long
    Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Sub DragAcceptFiles Lib "shell32.dll" (ByVal hwnd As Long, ByVal fAccept As Long)
    Private Declare Sub DragFinish Lib "shell32.dll" (ByVal HDROP As Long)
    Private Declare Function DragQueryFile Lib "shell32.dll" Alias "DragQueryFileA" (ByVal HDROP As Long, ByVal UINT As Long, ByVal lpStr As String, ByVal ch As Long) As Long
#End If


Private Sub TextBox1_Change()
    SelectControItemsByFilter ListBox1, TextBox1.Text
End Sub

Private Sub UserForm_Initialize()
    UserformOnTop Me
End Sub

Private Sub UserForm_Activate()

   #If VBA7 Then
        Dim hWnd As LongPtr, HDROP As LongPtr
    #Else
        Dim hWnd As Long, HDROP As Long
    #End If

    Const WM_DROPFILES = &H233
    Dim tMsg As MSG, sFileName As String * 256
    Dim lFilesCount As Long, i As Long


    Call WindowFromAccessibleObject(Me, hWnd)
    Call DragAcceptFiles(ListBox1.[_GethWnd], True)

    Do While GetMessage(tMsg, 0, 0, 0) And IsWindow(hWnd)
        If tMsg.message = WM_DROPFILES Then
            HDROP = tMsg.wParam
            lFilesCount = DragQueryFile(HDROP, &HFFFFFFFF, 0, 0)
            If lFilesCount Then
                For i = 0 To lFilesCount - 1
                    Dim CleanName As String
                    CleanName = Left(sFileName, DragQueryFile(HDROP, i, sFileName, Len(sFileName)))
                    If isFDU(CleanName) = "F" Then
                        ListBox1.AddItem
                        ListBox1.List(ListBox1.ListCount - 1, 0) = Mid(CleanName, InStrRev(CleanName, "\") + 1)
                        ListBox1.List(ListBox1.ListCount - 1, 1) = CleanName
                    Else
                        Dim element As Variant
                        Dim out As New Collection
                        FilesAndOrFoldersInFolderOrZip CleanName, oLogFolders, oLogFiles, oSearchInSubfolders, out
                        For Each element In out
                            ListBox1.AddItem element
                            ListBox1.List(ListBox1.ListCount - 1, 0) = Mid(element, InStrRev(element, "\") + 1)
                            ListBox1.List(ListBox1.ListCount - 1, 1) = element
                        Next
                    End If
                Next i
            End If
            Call DragFinish(HDROP)
        End If
        Call TranslateMessage(tMsg)
        Call DispatchMessage(tMsg)
    Loop
End Sub

Private Sub CommandButton1_Click()
    Dim element As Variant
    For Each element In ListboxSelectedValues(ListBox1)
        If CStr(element) Like "*.zip" Then
            UnzipToOwnFolder CStr(element), oDeleteExistingFolder, oDeleteZip
        End If
    Next
End Sub

Private Sub CommandButton2_Click()
    RemoveSelectedFromListbox ListBox1
End Sub

Sub RemoveSelectedFromListbox(lbox As MSForms.ListBox)
Dim i As Long
    Dim coll As New Collection
    Set coll = ListboxSelectedIndexes(lbox, False)
    For i = coll.Count To 1 Step -1
        lbox.RemoveItem coll(i)
    Next
End Sub

Private Sub info_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
uDEV.Show
End Sub

Rem file convert
Private Sub oExcelFiles_Click()
    WordOutput.Visible = False
    ExcelOutput.Visible = True
End Sub
Private Sub oWordFiles_Click()
    WordOutput.Visible = True
    WordOutput.Left = fFileType.Left
    ExcelOutput.Visible = False
End Sub

Private Sub Convert_Click()
    Dim element As Variant
    For Each element In ListboxSelectedValues(ListBox1)
        UnzipToOwnFolder CStr(element), oDeleteExistingFolder, oDeleteZip
    Next
End Sub

Sub convertFile(vPath As String)
    If oExcelFiles.Value = True Then
        If vPath Like "*.xl*" Then
            Select Case UCase(whichOption(Me.ExcelOutput, "OptionButton").Caption)
                Case "XLSB"
                    XLS_ConvertFileFormat vPath, xlExcel12, Me.oDelete
                Case "XLSM"
                    XLS_ConvertFileFormat vPath, xlOpenXMLWorkbookMacroEnabled, Me.oDelete
                Case "XLSX"
                    XLS_ConvertFileFormat vPath, xlWorkbookDefault, Me.oDelete
                Case "CSV"
                    XLS_ConvertFileFormat vPath, xlCSV, Me.oDelete
                Case "XLAM"
                    XLS_ConvertFileFormat vPath, xlOpenXMLAddIn, Me.oDelete
                Case "PDF"
                    ExcelToPDF vPath, cSeparateSheets.Value, True
            End Select
        End If
    Else
        If vPath Like "*.doc*" Then
            Select Case whichOption(Me.WordOutput, "OptionButton").Caption
                Case "DOCX"
                    Word_ConvertFileFormat vPath, wdFormatDocumentDefault, Me.oDelete
                Case "TXT"
                    Word_ConvertFileFormat vPath, wdFormatText, Me.oDelete
                Case "DOCM"
                    Word_ConvertFileFormat vPath, wdFormatXMLDocumentMacroEnabled, Me.oDelete
                Case "PDF"
                    Word_ConvertFileFormat vPath, wdFormatPDF, Me.oDelete
            End Select
        End If
    End If
End Sub






