Attribute VB_Name = "xlConverter"
Enum XlFileFormat
    'Ref: https://msdn.microsoft.com/en-us/vba/excel-vba/articles/xlfileformat-enumeration-excel
    xlAddIn = 18    'Microsoft Excel 97-2003 Add-In *.xla
    xlAddIn8 = 18    'Microsoft Excel 97-2003 Add-In *.xla
    xlCSV = 6    'CSV *.csv
    xlCSVMac = 22    'Macintosh CSV *.csv
    xlCSVMSDOS = 24    'MSDOS CSV *.csv
    xlCSVWindows = 23    'Windows CSV *.csv
    xlCurrentPlatformText = -4158    'Current Platform Text *.txt
    xlDBF2 = 7    'Dbase 2 format *.dbf
    xlDBF3 = 8    'Dbase 3 format *.dbf
    xlDBF4 = 11    'Dbase 4 format *.dbf
    xlDIF = 9    'Data Interchange format *.dif
    xlExcel12 = 50    'Excel Binary Workbook *.xlsb
    xlExcel2 = 16    'Excel version 2.0 (1987) *.xls
    xlExcel2FarEast = 27    'Excel version 2.0 far east (1987) *.xls
    xlExcel3 = 29    'Excel version 3.0 (1990) *.xls
    xlExcel4 = 33    'Excel version 4.0 (1992) *.xls
    xlExcel4Workbook = 35    'Excel version 4.0. Workbook format (1992) *.xlw
    xlExcel5 = 39    'Excel version 5.0 (1994) *.xls
    xlExcel7 = 39    'Excel 95 (version 7.0) *.xls
    xlExcel8 = 56    'Excel 97-2003 Workbook *.xls
    xlExcel9795 = 43    'Excel version 95 and 97 *.xls
    xlHtml = 44    'HTML format *.htm; *.html
    xlIntlAddIn = 26    'International Add-In No file extension
    xlIntlMacro = 25    'International Macro No file extension
    xlOpenDocumentSpreadsheet = 60    'OpenDocument Spreadsheet *.ods
    xlOpenXMLAddIn = 55    'Open XML Add-In *.xlam
    xlOpenXMLStrictWorkbook = 61    '(&;H3D) Strict Open XML file *.xlsx
    xlOpenXMLTemplate = 54    'Open XML Template *.xltx
    xlOpenXMLTemplateMacroEnabled = 53    'Open XML Template Macro Enabled *.xltm
    xlOpenXMLWorkbook = 51    'Open XML Workbook *.xlsx
    xlOpenXMLWorkbookMacroEnabled = 52    'Open XML Workbook Macro Enabled *.xlsm
    xlSYLK = 2    'Symbolic Link format *.slk
    xlTemplate = 17    'Excel Template format *.xlt
    xlTemplate8 = 17    ' Template 8 *.xlt
    xlTextMac = 19    'Macintosh Text *.txt
    xlTextMSDOS = 21    'MSDOS Text *.txt
    xlTextPrinter = 36    'Printer Text *.prn
    xlTextWindows = 20    'Windows Text *.txt
    xlUnicodeText = 42    'Unicode Text No file extension; *.txt
    xlWebArchive = 45    'Web Archive *.mht; *.mhtml
    xlWJ2WD1 = 14    'Japanese 1-2-3 *.wj2
    xlWJ3 = 40    'Japanese 1-2-3 *.wj3
    xlWJ3FJ3 = 41    'Japanese 1-2-3 format *.wj3
    xlWK1 = 5    'Lotus 1-2-3 format *.wk1
    xlWK1ALL = 31    'Lotus 1-2-3 format *.wk1
    xlWK1FMT = 30    'Lotus 1-2-3 format *.wk1
    xlWK3 = 15    'Lotus 1-2-3 format *.wk3
    xlWK3FM3 = 32    'Lotus 1-2-3 format *.wk3
    xlWK4 = 38    'Lotus 1-2-3 format *.wk4
    xlWKS = 4    'Lotus 1-2-3 format *.wks
    xlWorkbookDefault = 51    'Workbook default *.xlsx
    xlWorkbookNormal = -4143    'Workbook normal *.xls
    xlWorks2FarEast = 28    'Microsoft Works 2.0 far east format *.wks
    xlWQ1 = 34    'Quattro Pro format *.wq1
    xlXMLSpreadsheet = 46    'XML Spreadsheet *.xml
    
    xlTypePDF
End Enum
 
'---------------------------------------------------------------------------------------
' Procedure : XLS_ConvertFileFormat
' Author    : Daniel Pineault, CARDA Consultants Inc.
' Website   : http://www.cardaconsultants.com
' Purpose   : Converts an Excel compatible file format to another format
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
' Convert an xls file into a txt file and delete the xls once completed
'   Call XLS_ConvertFileFormat("C:TempTest.xls", xlTextWindows)
' Convert an xls file into a xlsx file and NOT delete the xls once completed
'   Call XLS_ConvertFileFormat("C:TempTest.xls",, False)
' Convert a csv file into a xlsx file and delete the xls once completed
'   Call XLS_ConvertFileFormat("C:TempTest.csv", xlWorkbookDefault, True)
'
' Revision History:
' Rev       Date(yyyy/mm/dd)        Description
' **************************************************************************************
' 1         2018-02-27              Initial Release
' 2         2020-12-31              Fixed typo xlDBF24 -> xlDBF4
'---------------------------------------------------------------------------------------
Function XLS_ConvertFileFormat(ByVal sOrigFile As String, _
                               Optional lNewFileFormat As XlFileFormat = xlOpenXMLWorkbook, _
                               Optional bDelOrigFile As Boolean = False) As Boolean
    '#Const EarlyBind = True 'Use Early Binding, Req. Reference Library
    #Const EarlyBind = False    'Use Late Binding
    #If EarlyBind = True Then
        'Early Binding Declarations
        Dim oExcel            As Excel.Application
        Dim oExcelWrkBk       As Excel.Workbook
    #Else
        'Late Binding Declaration/Constants
        Dim oExcel            As Object
        Dim oExcelWrkBk       As Object
    #End If
    Dim bExcelOpened          As Boolean
    Dim sOrigFileExt          As String
    Dim sNewXLSFileExt        As String
 
    'Determine the file extension associated with the requested file format
    'for properly renaming the output file
    Select Case lNewFileFormat
        Case xlAddIn, xlAddIn8
            sNewFileExt = ".xla"
        Case xlCSV, xlCSVMac, xlCSVMSDOS, xlCSVWindows
            sNewFileExt = ".csv"
        Case xlCurrentPlatformText, xlTextMac, xlTextMSDOS, xlTextWindows, xlUnicodeText
            sNewFileExt = ".txt"
        Case xlDBF2, xlDBF3, xlDBF4
            sNewFileExt = ".dbf"
        Case xlDIF
            sNewFileExt = ".dif"
        Case xlExcel12 = 50    'Excel Binary Workbook *.xlsb
            sNewFileExt = ".xlsb"
        Case xlExcel2, xlExcel2FarEast, xlExcel3, xlExcel4, xlExcel5, xlExcel7, _
             xlExcel8, xlExcel9795, xlWorkbookNormal
            sNewFileExt = ".xls"
        Case xlExcel4Workbook = 35    'Excel version 4.0. Workbook format (1992) *.xlw
            sNewFileExt = ".xlw"
        Case xlHtml = 44    'HTML format *.htm; *.html
            sNewFileExt = ".html"
        Case xlIntlAddIn, xlIntlMacro
            sNewFileExt = ""
        Case xlOpenDocumentSpreadsheet    'OpenDocument Spreadsheet *.ods
            sNewFileExt = ".ods"
        Case xlOpenXMLAddIn    'Open XML Add-In *.xlam
            sNewFileExt = ".xlam"
        Case xlOpenXMLStrictWorkbook, xlOpenXMLWorkbook, xlWorkbookDefault = 51
            sNewFileExt = ".xlsx"
        Case xlOpenXMLTemplate    'Open XML Template *.xltx
            sNewFileExt = ".xltx"
        Case xlOpenXMLTemplateMacroEnabled     'Open XML Template Macro Enabled *.xltm
            sNewFileExt = ".xltm"
        Case xlOpenXMLWorkbookMacroEnabled     'Open XML Workbook Macro Enabled *.xlsm
            sNewFileExt = ".xlsm"
        Case xlSYLK     'Symbolic Link format *.slk
            sNewFileExt = ".slk"
        Case xlTemplate, xlTemplate8    ' Template 8 *.xlt
            sNewFileExt = ".xlt"
        Case xlTextPrinter        'Printer Text *.prn
            sNewFileExt = ".prn"
        Case xlWebArchive         'Web Archive *.mht; *.mhtml
            sNewFileExt = ".mhtml"
        Case xlWJ2WD1        'Japanese 1-2-3 *.wj2
            sNewFileExt = ".wj2"
        Case xlWJ3, xlWJ3FJ3    'Japanese 1-2-3 format *.wj3
            sNewFileExt = ".wj3"
        Case xlWK1, xlWK1ALL, xlWK1FMT   'Lotus 1-2-3 format *.wk1
            sNewFileExt = ".wk1"
        Case xlWK3, xlWK3FM3   'Lotus 1-2-3 format *.wk3
            sNewFileExt = ".wk3"
        Case xlWK4       'Lotus 1-2-3 format *.wk4
            sNewFileExt = ".wk4"
        Case xlWKS, xlWorks2FarEast      'Lotus 1-2-3 format *.wks
            sNewFileExt = ".wks"
        Case xlWQ1       'Quattro Pro format *.wq1
            sNewFileExt = ".wq1"
        Case xlXMLSpreadsheet       'XML Spreadsheet *.xml
            sNewFileExt = ".xml"
            
        Case xlTypePDF
            sNewFileExt = ".pdf"
            'todo
            'ExcelToPDF
            'exit function
    End Select
 
    'Determine the original file's extension for properly renaming the output file
    sOrigFileExt = "." & Right(sOrigFile, Len(sOrigFile) - InStrRev(sOrigFile, "."))
 
    'Start Excel
    On Error Resume Next
    Set oExcel = GetObject(, "Excel.Application")          'Bind to existing instance of Excel
    If Err.Number <> 0 Then          'Could not get instance of Excel, so create a new one
        Err.Clear
        On Error GoTo Error_Handler
        Set oExcel = CreateObject("Excel.Application")
    Else          'Excel was already running
        bExcelOpened = True
    End If
    On Error GoTo Error_Handler
 
    oExcel.ScreenUpdating = False
    oExcel.Visible = False         'Keep Excel hidden until we are done with our manipulation
    Set oExcelWrkBk = oExcel.Workbooks.Open(sOrigFile)    'Open the original file
    'Save it as the requested new file format
    oExcelWrkBk.SaveAs Replace(sOrigFile, sOrigFileExt, sNewFileExt), lNewFileFormat, , , , False
    XLS_ConvertFileFormat = True    'Report back that we managed to save the file in the new format
    oExcelWrkBk.Close False    'Close the workbook
    If bExcelOpened = False Then
        oExcel.Quit    'Quit Excel only if we started it
    Else
        oExcel.ScreenUpdating = True
        oExcel.Visible = True
    End If
 
    'If bDelOrigFile = True Then Kill (sOrigFile)    'Delete the original file if requested
    If bDelOrigFile = True Then Recycle (sOrigFile)
    
Error_Handler_Exit:
    On Error Resume Next
    Set oExcelWrkBk = Nothing
    Set oExcel = Nothing
    Exit Function
 
Error_Handler:
    MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Error Source: XLS_ConvertFileFormat" & vbCrLf & _
           "Error Description: " & Err.Description & _
           Switch(Erl = 0, "", Erl <> 0, vbCrLf & "Line No: " & Erl) _
           , vbOKOnly + vbCritical, "An Error has Occurred!"
    oExcel.ScreenUpdating = True
    oExcel.Visible = True         'Make excel visible to the user
    Resume Error_Handler_Exit
End Function



Sub ExcelToPDF(fileFullPath As String, SeparateSheets As Boolean, CloseFile As Boolean)

Dim wb As Workbook
    Set wb = Workbooks.Open(fileFullPath)
Dim ws As Worksheet

If SeparateSheets = False Then
    wb.ExportAsFixedFormat xlTypePDF, _
    VBA.Replace(fileFullPath, Right(fileFullPath, Len(fileFullPath) - InStrRev(fileFullPath, ".") + 1), ".pdf")
    If CloseFile = True Then wb.Close False
Else
    For Each ws In wb
    ws.ExportAsFixedFormat xlTypePDF, wb.path & "\" & ws.Name & ".pdf"
    Next ws
End If
MsgBox "Process Completed"
End Sub
