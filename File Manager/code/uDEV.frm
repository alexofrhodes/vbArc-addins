VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uDEV 
   Caption         =   "vbArc ~ Anastasiou Alex"
   ClientHeight    =   2580
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4044
   OleObjectBlob   =   "uDEV.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uDEV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub LFaceBook_Click()
FollowLink ("https://www.facebook.com/VBA-Code-Archive-110295994460212")
End Sub

Private Sub LGitHub_Click()
FollowLink ("https://github.com/alexofrhodes")
End Sub

Private Sub LYouTube_Click()
FollowLink ("https://bit.ly/2QT4wFe")
End Sub

Private Sub LBuyMeACoffee_Click()
FollowLink ("http://paypal.me/alexofrhodes")
End Sub

Private Sub LEmail_Click()
If OutlookCheck = True Then
    MailDev
Else
    Dim out As String
    out = "anastasioualex@gmail.com"
    CLIP out
    MsgBox ("Outlook not found" & Chr(10) & _
    "DEV's email address" & vbNewLine & out & vbNewLine & "copied to clipboard")
End If
End Sub

Sub MailDev()
    'For Tips see: http://www.rondebruin.nl/win/winmail/Outlook/tips.htm
    'Working in Office 2000-2016
    Dim OutApp As Object
    Dim OutMail As Object
    Dim strBody As String
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    '    strbody = "Hi there" & vbNewLine & vbNewLine & _
    "This is line 1" & vbNewLine & _
    "This is line 2" & vbNewLine & _
    "This is line 3" & vbNewLine & _
    "This is line 4"
    On Error Resume Next
    With OutMail
        .To = "anastasioualex@gmail.com"
        .CC = vbNullString
        .BCC = vbNullString
        .Subject = "DEV REQUEST OR FEEDBACK FOR -CODE ARCHIVE-"
        .body = strBody
        'You can add a file like this
        '.Attachments.Add ("C:\test.txt")
        '.Send
        .Display
    End With
    On Error GoTo 0
    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub

Sub FollowLink(folderPath As String)
        Dim oShell As Object
        Dim Wnd As Object
        Set oShell = CreateObject("Shell.Application")
        For Each Wnd In oShell.Windows
            If Wnd.Name = "File Explorer" Then
               If Wnd.Document.Folder.Self.path = folderPath Then Exit Sub
            End If
        Next Wnd
        Application.ThisWorkbook.FollowHyperlink Address:=folderPath, NewWindow:=True
End Sub
