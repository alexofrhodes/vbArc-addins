Attribute VB_Name = "wordConverter"
 Enum WdSaveFormat
    'Ref: https://msdn.microsoft.com/en-us/vba/word-vba/articles/wdsaveformat-enumeration-word
    wdFormatDocument = 0    'Microsoft Office Word 97 - 2003 binary file format.
    wdFormatDOSText = 4    'Microsoft DOS text format.  *.txt
    wdFormatDOSTextLineBreaks = 5    'Microsoft DOS text with line breaks preserved.  *.txt
    wdFormatEncodedText = 7    'Encoded text format.  *.txt
    wdFormatFilteredHTML = 10    'Filtered HTML format.
    wdFormatFlatXML = 19    'Open XML file format saved as a single XML file.
'    wdFormatFlatXML = 20    'Open XML file format with macros enabled saved as a single XML file.
    wdFormatFlatXMLTemplate = 21    'Open XML template format saved as a XML single file.
    wdFormatFlatXMLTemplateMacroEnabled = 22    'Open XML template format with macros enabled saved as a single XML file.
    wdFormatOpenDocumentText = 23    'OpenDocument Text format. *.odt
    wdFormatHTML = 8    'Standard HTML format. *.html
    wdFormatRTF = 6    'Rich text format (RTF). *.rtf
    wdFormatStrictOpenXMLDocument = 24    'Strict Open XML document format.
    wdFormatTemplate = 1    'Word template format.
    wdFormatText = 2    'Microsoft Windows text format. *.txt
    wdFormatTextLineBreaks = 3    'Windows text format with line breaks preserved. *.txt
    wdFormatUnicodeText = 7    'Unicode text format. *.txt
    wdFormatWebArchive = 9    'Web archive format.
    wdFormatXML = 11    'Extensible Markup Language (XML) format. *.xml
    wdFormatDocument97 = 0    'Microsoft Word 97 document format. *.doc
    wdFormatDocumentDefault = 16    'Word default document file format. For Word, this is the DOCX format. *.docx
    wdFormatPDF = 17    'PDF format. *.pdf
    wdFormatTemplate97 = 1    'Word 97 template format.
    wdFormatXMLDocument = 12    'XML document format.
    wdFormatXMLDocumentMacroEnabled = 13    'XML document format with macros enabled.
    wdFormatXMLTemplate = 14    'XML template format.
    wdFormatXMLTemplateMacroEnabled = 15    'XML template format with macros enabled.
    wdFormatXPS = 18    'XPS format. *.xps
End Enum

'---------------------------------------------------------------------------------------
' Procedure : Word_ConvertFileFormat
' Author    : Daniel Pineault, CARDA Consultants Inc.
' Website   : http://www.cardaconsultants.com
' Purpose   : Converts a Word compatible file format to another format
' Copyright : The following is release as Attribution-ShareAlike 4.0 International
'             (CC BY-SA 4.0) - https://creativecommons.org/licenses/by-sa/4.0/
' Req'd Refs: Uses Late Binding, so none required
'
' Input Variables:
' ~~~~~~~~~~~~~~~~
' sOrigFile     : String - Original file path, name and extension to be converted
' lNewFileFormat: New File format to save the original file as
' bDelOrigFile  : True/False - Should the original file be deleted after the conversion
'
' Usage:
' ~~~~~~
' Convert a doc file into a docx file but retain the original copy
'   Call Word_ConvertFileFormat("C:\Users\Daniel\Documents\Resume.doc", wdFormatPDF)
' Convert a doc file into a docx file and delete the original doc once converted
'   Call Word_ConvertFileFormat("C:\Users\Daniel\Documents\Resume.doc", wdFormatPDF, True)
'
' Revision History:
' Rev       Date(yyyy/mm/dd)        Description
' **************************************************************************************
' 1         2018-02-27              Initial Release
'---------------------------------------------------------------------------------------
Function Word_ConvertFileFormat(ByVal sOrigFile As String, _
                                Optional lNewFileFormat As WdSaveFormat = wdFormatDocumentDefault, _
                                Optional bDelOrigFile As Boolean = False) As Boolean
    '#Const EarlyBind = True 'Use Early Binding, Req. Reference Library
    #Const EarlyBind = False    'Use Late Binding
    #If EarlyBind = True Then
        'Early Binding Declarations
        Dim oWord             As Word.Application
        Dim oDoc              As Word.Document
    #Else
        'Late Binding Declaration/Constants
        Dim oWord             As Object
        Dim oDoc              As Object
    #End If
    Dim bWordOpened           As Boolean
    Dim sOrigFileExt          As String
    Dim sNewFileExt           As String
 
    'Determine the file extension associated with the requested file format
    'for properly renaming the output file
    Select Case lNewFileFormat
        Case wdFormatDocument
            sNewFileExt = "."
        Case wdFormatDOSText, wdFormatDOSTextLineBreaks, wdFormatEncodedText, wdFormatOpenDocumentText, wdFormatText, wdFormatTextLineBreaks, wdFormatUnicodeText
            sNewFileExt = ".txt"
        Case wdFormatFilteredHTML, wdFormatHTML
            sNewFileExt = ".html"
        Case wdFormatFlatXML, wdFormatXML, wdFormatXMLDocument
            sNewFileExt = ".xml"
        Case wdFormatFlatXMLTemplate
            sNewFileExt = "."
        Case wdFormatFlatXMLTemplateMacroEnabled
            sNewFileExt = "."
        Case wdFormatRTF
            sNewFileExt = ".rtf"
        Case wdFormatStrictOpenXMLDocument
            sNewFileExt = "."
        Case wdFormatTemplate
            sNewFileExt = "."
        Case wdFormatWebArchive
            sNewFileExt = "."
        Case wdFormatDocument97
            sNewFileExt = ".doc"
        Case wdFormatDocumentDefault
            sNewFileExt = ".docx"
        Case wdFormatPDF
            sNewFileExt = ".pdf"
        Case wdFormatTemplate97
            sNewFileExt = "."
        Case wdFormatXMLDocumentMacroEnabled
            sNewFileExt = ".docm"
        Case wdFormatXMLTemplate
            sNewFileExt = ".doct"
        Case wdFormatXMLTemplateMacroEnabled
            sNewFileExt = "."
        Case wdFormatXPS
            sNewFileExt = ".xps"
    End Select
 
    'Determine the original file's extension for properly renaming the output file
    sOrigFileExt = "." & Right(sOrigFile, Len(sOrigFile) - InStrRev(sOrigFile, "."))
 
    'Start Excel
    On Error Resume Next
    Set oWord = GetObject(, "Word.Application")            'Bind to existing instance of Word
    If Err.Number <> 0 Then            'Could not get instance of Word, so create a new one
        Err.Clear
        On Error GoTo Error_Handler
        Set oWord = CreateObject("Word.Application")
    Else            'Word was already running
        bWordOpened = True
    End If
    On Error GoTo Error_Handler
 
    oWord.Visible = False           'Keep Word hidden until we are done with our manipulation
    Set oDoc = oWord.Documents.Open(sOrigFile)      'Open the original file
    'Save it as the requested new file format
    oDoc.SaveAs2 Replace(sOrigFile, sOrigFileExt, sNewFileExt), lNewFileFormat
    Word_ConvertFileFormat = True      'Report back that we managed to save the file in the new format
    oDoc.Close False      'Close the document
    If bWordOpened = False Then
        oWord.Quit      'Quit Word only if we started it
    Else
        oWord.Visible = True 'Since it was already open, ensure it is visible
    End If
 
    'If bDelOrigFile = True Then Kill (sOrigFile)      'Delete the original file if requested
    If bDelOrigFile = True Then Recycle (sOrigFile)
    
Error_Handler_Exit:
    On Error Resume Next
    Set oDoc = Nothing
    Set oWord = Nothing
    Exit Function
 
Error_Handler:
    MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Error Source: XLS_ConvertFileFormat" & vbCrLf & _
           "Error Description: " & Err.Description & _
           Switch(Erl = 0, "", Erl <> 0, vbCrLf & "Line No: " & Erl) _
           , vbOKOnly + vbCritical, "An Error has Occurred!"
    oWord.Visible = True           'Make excel visible to the user
    Resume Error_Handler_Exit
End Function
