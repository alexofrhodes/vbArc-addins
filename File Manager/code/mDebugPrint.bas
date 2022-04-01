Attribute VB_Name = "mDebugPrint"
Sub dp(var As Variant)
    Dim element As Variant
    Select Case TypeName(var)
        Case Is = "String", "Long", "Integer", "Boolean"
            Debug.Print var
            Rem todo How to handle multidimensional array?
        Case Is = "Variant()", "String()", "Long()", "Integer()"
            If ArrayDimensions(var) = 1 Then
                Dim i As Long
                For i = LBound(var) To UBound(var)
                    Debug.Print var(i)
                Next i
            ElseIf ArrayDimensions(var) > 1 Then
                DPH var
            End If
        Case Is = "Collection"
            For Each element In var
                dp element
            Next element
        Case Else
    End Select
End Sub

Function DpHeader(str As Variant, Optional lvl As Integer = 1, Optional Character As String = "'", _
                  Optional Top As Boolean, Optional bottom As Boolean) As String
    If lvl < 1 Then lvl = 1
    If Character = "" Then Character = "'"
    Dim indentation As Integer
    indentation = (lvl * 4) - 4 + 1
    Dim QUOTE As String: QUOTE = "'"
    Dim s As String
    Dim element As Variant
    If Top = True Then s = vbNewLine & QUOTE & String(indentation + LargestLength(str), Character) & vbNewLine
    If TypeName(str) <> "String" Then
        For Each element In str
            s = s & QUOTE & String(indentation, Character) & element & vbNewLine
        Next
    Else
        s = s & QUOTE & String(indentation, Character) & str
    End If
    If bottom = True Then s = s & QUOTE & String(indentation + LargestLength(str), Character)
    DpHeader = s
End Function

Public Function Len2( _
       ByVal val As Variant) _
        As Integer
    If IsArray(val) And Right(TypeName(val), 2) = "()" Then
        Len2 = UBound(val) - LBound(val) + 1
    ElseIf TypeName(val) = "String" Then
        Len2 = Len(val)
    ElseIf IsNumeric(val) Then
        Len2 = Len(CStr(val))
    Else
        Len2 = val.Count
    End If
End Function

Function LargestLength(Optional myObj) As Long
    LargestLength = 0
    Dim element As Variant
    If IsMissing(myObj) Then
        If TypeName(Selection) = "Range" Then
            Set myObj = Selection
        Else
            Exit Function
        End If
    End If
    Select Case TypeName(myObj)
        Case Is = "String"
            LargestLength = Len(myObj)
        Case "Collection"
            For Each element In myObj
                If Len(element) > LargestLength Then LargestLength = Len(element)
            Next element
        Case "Variant", "Variant()", "String()"
            For element = LBound(myObj) To UBound(myObj)
                If Len(myObj(element)) > LargestLength Then LargestLength = Len(myObj(element))
            Next
        Case Else
    End Select
End Function

Sub DPH(ByVal Hairetu, Optional HyoujiMaxNagasa%, Optional HairetuName$)
    Call DebugPrintHairetu(Hairetu, HyoujiMaxNagasa, HairetuName)
End Sub

Sub DebugPrintHairetu(ByVal Hairetu, Optional HyoujiMaxNagasa%, Optional HairetuName$)
    Dim i&, j&, k&, m&, n&
    Dim TateMin&, TateMax&, YokoMin&, YokoMax&
    Dim WithTableHairetu
    Dim NagasaList, MaxNagasaList
    Dim NagasaOnajiList
    Dim OutputList
    Const SikiriMoji$ = "|"
    Dim Jigen2%
    On Error Resume Next
    Jigen2 = UBound(Hairetu, 2)
    On Error GoTo 0
    If Jigen2 = 0 Then
        Hairetu = Application.Transpose(Hairetu)
    End If
    TateMin = LBound(Hairetu, 1)
    TateMax = UBound(Hairetu, 1)
    YokoMin = LBound(Hairetu, 2)
    YokoMax = UBound(Hairetu, 2)
    ReDim WithTableHairetu(1 To TateMax - TateMin + 1 + 1, 1 To YokoMax - YokoMin + 1 + 1)
    For i = 1 To TateMax - TateMin + 1
        WithTableHairetu(i + 1, 1) = TateMin + i - 1
            For j = 1 To YokoMax - YokoMin + 1
                WithTableHairetu(1, j + 1) = YokoMin + j - 1
                    WithTableHairetu(i + 1, j + 1) = Hairetu(i - 1 + TateMin, j - 1 + YokoMin)
                    Next j
                Next i
                n = UBound(WithTableHairetu, 1)
                m = UBound(WithTableHairetu, 2)
                ReDim NagasaList(1 To n, 1 To m)
                ReDim MaxNagasaList(1 To m)
                Dim TmpStr$
                For j = 1 To m
                    For i = 1 To n
                        If j > 1 And HyoujiMaxNagasa <> 0 Then
                            TmpStr = WithTableHairetu(i, j)
                            WithTableHairetu(i, j) = ShortenToByteCharacters(TmpStr, HyoujiMaxNagasa)
                            End If
NagasaList(i, j) = LenB(StrConv(WithTableHairetu(i, j), vbFromUnicode))
                            MaxNagasaList(j) = WorksheetFunction.Max(MaxNagasaList(j), NagasaList(i, j))
                        Next i
                    Next j
                    ReDim NagasaOnajiList(1 To n, 1 To m)
                    Dim TmpMaxNagasa&
                    For j = 1 To m
                        TmpMaxNagasa = MaxNagasaList(j)
                        For i = 1 To n
                            NagasaOnajiList(i, j) = WithTableHairetu(i, j) & WorksheetFunction.Rept(" ", TmpMaxNagasa - NagasaList(i, j))
                        Next i
                    Next j
                    ReDim OutputList(1 To n)
                    For i = 1 To n
                        For j = 1 To m
                            If j = 1 Then
                                OutputList(i) = NagasaOnajiList(i, j)
                            Else
                                OutputList(i) = OutputList(i) & SikiriMoji & NagasaOnajiList(i, j)
                            End If
                        Next j
                    Next i
                    Debug.Print HairetuName
                    For i = 1 To n
                        Debug.Print OutputList(i)
                    Next i
                End Sub

Function ShortenToByteCharacters(Mojiretu$, ByteNum%)
    Dim OriginByte%
    Dim Output
    OriginByte = LenB(StrConv(Mojiretu, vbFromUnicode))
    If OriginByte <= ByteNum Then
        Output = Mojiretu
    Else
        Dim RuikeiByteList, BunkaiMojiretu
        RuikeiByteList = CalculateByteCharacters(Mojiretu)
        BunkaiMojiretu = TextDecomposition(Mojiretu)
        Dim AddMoji$
        AddMoji = "."
        Dim i&, n&
        n = Len(Mojiretu)
        For i = 1 To n
            If RuikeiByteList(i) < ByteNum Then
                Output = Output & BunkaiMojiretu(i)
            ElseIf RuikeiByteList(i) = ByteNum Then
                If LenB(StrConv(BunkaiMojiretu(i), vbFromUnicode)) = 1 Then
                    Output = Output & AddMoji
                Else
                    Output = Output & AddMoji & AddMoji
                End If
                Exit For
            ElseIf RuikeiByteList(i) > ByteNum Then
                Output = Output & AddMoji
                Exit For
            End If
        Next i
    End If
    ShortenToByteCharacters = Output
End Function

Function CalculateByteCharacters(Mojiretu$)
    Dim MojiKosu%
    MojiKosu = Len(Mojiretu)
    Dim Output
    ReDim Output(1 To MojiKosu)
    Dim i&
    Dim TmpMoji$
    For i = 1 To MojiKosu
        TmpMoji = Mid(Mojiretu, i, 1)
        If i = 1 Then
            Output(i) = LenB(StrConv(TmpMoji, vbFromUnicode))
        Else
            Output(i) = LenB(StrConv(TmpMoji, vbFromUnicode)) + Output(i - 1)
        End If
    Next i
    CalculateByteCharacters = Output
End Function

Function TextDecomposition(Mojiretu$)
    Dim i&, n&
    Dim Output
    n = Len(Mojiretu)
    ReDim Output(1 To n)
    For i = 1 To n
        Output(i) = Mid(Mojiretu, i, 1)
    Next i
    TextDecomposition = Output
End Function

Function ArrayDimensions(ByVal vArray As Variant) As Long
    Dim dimnum As Long
    On Error GoTo FinalDimension
    For dimnum = 1 To 60000
        ErrorCheck = LBound(vArray, dimnum)
    Next
FinalDimension:
    ArrayDimensions = dimnum - 1
End Function


