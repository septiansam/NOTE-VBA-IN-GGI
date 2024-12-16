
Sub PROCESS()

Dim starttime As Double
Dim elapsedtime As Double
Dim ws As Worksheet
Dim rng As Range
Dim delimiters As String
Dim i As Long, j As Long, k As Long
Dim currentChar As String
Dim rgPengali As Range, pengali As Long
Dim sumSize As Long
starttime = Timer

' 12-26-2023 'S.A.M
Dim cell As Range, lc As Long, lr As Long
Sheets("MA").Select
lr = Range("H" & Rows.Count).End(xlUp).Row
Set rng = Range("H30:H" & lr)
For Each cell In rng
    If cell.MergeCells Then
        ' Nonaktifkan Merge & Center
        cell.MergeCells = False
    End If
Next cell
' Selesai

Sheets("MA").Copy After:=Sheets(4)
Sheets("MA (2)").Name = "MA use"

Sheets("MA use").Select
Range("M" & Rows.Count).End(xlUp).Select
SUM1 = ActiveCell.Row
Range("B" & Rows.Count).End(xlUp).Select
SUM2 = ActiveCell.Row
Range("H" & Rows.Count).End(xlUp).Select
SUM3 = ActiveCell.Row

If SUM1 > SUM2 And SUM1 > SUM3 Then
    SUMALL = SUM1
ElseIf SUM2 > SUM1 And SUM2 > SUM3 Then
    SUMALL = SUM2
ElseIf SUM3 > SUM1 And SUM3 > SUM2 Then
    SUMALL = SUM3
ElseIf SUM1 = SUM2 = SUM3 Then
    SUMALL = SUM1
End If

SUMALL = Cells.Find(What:="*", lookat:=xlPart, LookIn:=xlFormulas, Searchorder:=xlByRows, searchdirection:=xlPrevious).Row
    
Range("H30:H" & SUMALL).Replace What:="/", Replacement:="="

Range("A" & Rows.Count).End(xlUp).Select
Range(Selection, "A30").Select
Selection.NumberFormat = "[$-en-US]d-mmm-yy;@"
Selection.HorizontalAlignment = xlCenter
Selection.VerticalAlignment = xlBottom

For i = 30 To SUMALL
If Cells(i, 7) <> "" And Cells(1 + i, 7) = "" Then
Cells(i, 17) = i
End If
Next i

Sheets.Add(After:=Sheets(Sheets.Count)).Name = "TES"

Sheets("MA use").Select
Range("H" & Rows.Count).End(xlUp).Select
Range(Selection, "H30").Copy
Sheets("TES").Select
Range("A1").PasteSpecial xlPasteAll

Dim LastRowIndex As Integer
Dim RowIndex As Integer
Dim UsedRng As Range
Set UsedRng = ActiveSheet.UsedRange
    LastRowIndex = UsedRng.Row - 1 + UsedRng.Rows.Count
Application.ScreenUpdating = False
For RowIndex = LastRowIndex To 1 Step -1
If Application.CountA(Rows(RowIndex)) = 0 Then
Rows(RowIndex).Delete
End If
Next RowIndex
Application.ScreenUpdating = True

Range("A" & Rows.Count).End(xlUp).Select
sumSize = ActiveCell.Row / 2

Sheets.Add(After:=Sheets(Sheets.Count)).Name = "TES1"

Set ws = ThisWorkbook.Sheets("TES1")
For i = 1 To sumSize
    Sheets("TES").Select
    Cells(2 * i, 1).Copy
'    Sheets("TES1").Select
'    Range("A1").PasteSpecial xlPasteValues
    
'[*]__SEPTIAN <-- 30-JULI-2024 -->
    ws.Activate
    Range("A1").PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
    Set rng = ws.Range("A1")
    
    ' Daftar delimiter yang umum digunakan
    delimiters = "= ,.;:|/\~!@#$%^&*()_-+{}[]<>?`'"""
    
    ' Loop untuk mengganti semua delimiter dengan satu pembatas '|'
    For j = 1 To Len(delimiters)
        currentChar = Mid(delimiters, j, 1)
        rng.Value = Replace(rng.Value, currentChar, "|")
    Next j
    
    Data = rng.Value
    delimiter = "|"

    ' Ganti semua urutan delimiter berturut-turut dengan satu delimiter
    Do While InStr(Data, delimiter & delimiter) > 0
        Data = Replace(Data, delimiter & delimiter, delimiter)
    Loop
    
    rng.Value = Data

    ' Pisahkan teks berdasarkan pembatas '|'
    rng.TextToColumns Destination:=rng, _
                      DataType:=xlDelimited, _
                      TextQualifier:=xlDoubleQuote, _
                      ConsecutiveDelimiter:=True, _
                      Other:=True, _
                      OtherChar:="|", _
                      FieldInfo:=Array(1, 1), _
                      TrailingMinusNumbers:=True

'[*]__DONE........................
    
'    Range("A1").TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, Other:=True, OtherChar:="+", FieldInfo:=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
'    If Range("B1") <> "" Then
'        Range("B1").Copy
'        Range("A3").PasteSpecial xlPasteValues
'        Range("B1").ClearContents
'    End If
'    Range("A1").TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, Other:=True, OtherChar:="-", FieldInfo:=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
'
'    If Range("B1") <> "" Then
'        Range("A3") = -Range("B1").Value
'        Range("B1").ClearContents
'    End If
'
'    Range("A1").TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, Other:=True, OtherChar:="#", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1)), TrailingMinusNumbers:=True
'    Range("B1").Copy
'    Range("A4").PasteSpecial xlPasteValues
'    Range("B1").ClearContents
'    Range("A1").TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, Other:=True, OtherChar:="=", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1)), TrailingMinusNumbers:=True
'    Range("A1").Select
'    'SUMSIZE2 = Range(Selection, Selection.End(xlToRight)).Count
'    SUMSIZE2 = Application.WorksheetFunction.CountA(Rows(1))
'
'    For j = 2 To SUMSIZE2
'        If Mid(Cells(1, j), 2, 1) = " " Then
'            Cells(2, j) = Left(Cells(1, j), 1)
'        ElseIf Right(Cells(1, j), 1) <> ")" Then
'            Cells(2, j) = Left(Cells(1, j), 2)
'        End If
'    Next j
'
'    If Range("A4") <> "" Then
'        Range("B4") = Mid(Range("A4"), 2, Len(Range("A4")) - 1)
'        Range("A6") = ((Range("B2").Value + Range("C2").Value + Range("D2").Value + Range("E2").Value + Range("F2").Value + Range("G2").Value + Range("H2").Value + Range("I2").Value + Range("J2").Value + Range("K2").Value + Range("L2").Value + Range("M2").Value + Range("N2").Value + Range("O2").Value + Range("P2").Value + Range("Q2").Value + Range("R2").Value + Range("S2").Value + Range("T2").Value + Range("U2").Value + Range("V2").Value + Range("W2").Value + Range("X2").Value + Range("Y2").Value + Range("Z2").Value + Range("AA2").Value + Range("AB2").Value + Range("AC2").Value + Range("AD2").Value + Range("AE2").Value) * Range("B4").Value) + Range("A3").Value
'    ElseIf Range("A4").Value = "" Then
'        Range("A6") = (Range("B2").Value + Range("C2").Value + Range("D2").Value + Range("E2").Value + Range("F2").Value + Range("G2").Value + Range("H2").Value + Range("I2").Value + Range("J2").Value + Range("K2").Value + Range("L2").Value + Range("M2").Value + Range("N2").Value + Range("O2").Value + Range("P2").Value + Range("Q2").Value + Range("R2").Value + Range("S2").Value + Range("T2").Value + Range("U2").Value + Range("V2").Value + Range("W2").Value + Range("X2").Value + Range("Y2").Value + Range("Z2").Value + Range("AA2").Value + Range("AB2").Value + Range("AC2").Value + Range("AD2").Value + Range("AE2").Value) + Range("A3").Value
'    End If
'
'    Sheets("TES").Cells(2 * i, 2) = Sheets("TES1").Range("A6").Value

'[*]__SEPTIAN <-- 30-JULI-2024 -->
    lc = Cells(1, Columns.Count).End(xlToLeft).Column
    
    ''Cari Pengali
    
    For j = 2 To lc
        If Not IsNumeric(Cells(1, j)) And Mid(Cells(1, j), 1, 1) = "X" And IsNumeric(Right(Cells(1, j), 1)) Then
            Set rgPengali = Cells(1, j)
            Exit For
        End If
    Next j
    If Not rgPengali Is Nothing Then
        rgPengali.Replace " ", ""
        pengali = Right(rgPengali.Value, Len(rgPengali) - 1)
    Else
        pengali = 1
    End If
 
    ''Cari Jumlah
    sumSize = 0
    For j = 2 To lc
        If IsNumeric(Cells(1, j)) Then
            sumSize = sumSize + Cells(1, j)
        End If
    Next j
    
    sumSize = sumSize * pengali
    Sheets("TES").Cells(2 * i, 2) = sumSize

'[*]__DONE........................
    
    Sheets("TES1").Select
    Cells.ClearContents
    Set rgPengali = Nothing

Next i

Sheets("MA use").Select
Range("I30").FormulaR1C1 = "=VLOOKUP(R[1]C[-1],TES!C[-8]:C[-7],2,0)"
Range("I30").Copy
Range("H" & Rows.Count).End(xlUp).Offset(0, 1).Select
Range(Selection, "I30").Select
ActiveSheet.Paste
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Selection.Replace What:="#N/A", Replacement:=""
Selection.Replace "0", "", xlWhole
Application.CutCopyMode = False

Sheets("TES").Cells.Delete

Sheets("MA use").Select
Range("Q" & Rows.Count).End(xlUp).Select
Range(Selection, "Q30").Copy
Sheets("TES").Select
Range("A1").PasteSpecial xlPasteAll

Sheets("MA use").Select
If Range("J" & Rows.Count).End(xlUp) <> "LENGTH" Then
    Range("J" & Rows.Count).End(xlUp).Select
    Range(Selection, "J30").Copy
    Sheets("TES").Select
    Range("B1").PasteSpecial xlPasteAll
End If

Sheets("TES").Select
Application.ScreenUpdating = False
For RowIndex = LastRowIndex To 1 Step -1
    If Application.CountA(Rows(RowIndex)) = 0 Then
        Rows(RowIndex).Delete
    End If
Next RowIndex
Application.ScreenUpdating = True

Range("A1").CurrentRegion.Copy
Range("A1").PasteSpecial xlPasteValues
Application.CutCopyMode = False
Range("a1").Select

Range("A" & Rows.Count).End(xlUp).Select
SUMNO = ActiveCell.Row

Sheets("MA use").Select
Range("Q" & SUMALL + 1) = "QWERTY"

Sheets("FORMAT").Visible = True

Sheets("TES").Select
Range("A1").CurrentRegion.Copy
Range("A1").PasteSpecial xlPasteValues
Application.CutCopyMode = False
Range("a1").Select

For i = 1 To SUMNO
    Sheets("TES").Select
    Range("A1").CurrentRegion.Copy
    Range("A1").PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    Range("a1").Select
    Range("B:B").Replace What:="#REF!", Replacement:=""
    
    If Cells(i, 2) = "" Then
        Application.DisplayAlerts = False
        If Evaluate("isref('" & "A" & i & "'!A1)") Then Sheets("A" & i).Delete
        
        Sheets.Add(After:=Sheets(Sheets.Count)).Name = "A" & i
        
        Sheets("FORMAT").Select
        Cells.Copy
        Sheets("A" & i).Select
        Range("A1").PasteSpecial xlPasteAll
        'Range("J3").Formula = ""
        
        For j = 30 To SUMALL
        
            A = Sheets("TES").Cells(i, 1).Value
            B = Sheets("MA use").Cells(j, 17).Value
            
            If A = B Then
                Sheets("MA").Select
                Cells(j, 17) = "A" & i
                Cells(j, 17).Select
                ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:="'" & "A" & i & "'" & "!" & "A" & i
                Sheets("MA use").Select
                Cells(j, 17).Select
                Selection.End(xlDown).Offset(-1, -1).Select
                Range(Selection, Cells(j, 17).Offset(0, -1)).Copy
                Sheets("A" & i).Select
                Range("M3").PasteSpecial xlPasteAll
                
                Sheets("MA use").Select
                Cells(j, 17).Select
                Selection.End(xlDown).Offset(-1, -30).Select
                Range(Selection, Cells(j, 17).Offset(0, -8)).Copy
                Sheets("A" & i).Select
                Range("A3").PasteSpecial xlPasteAll
                
                Sheets("MA use").Select
                Cells(j, 17).Offset(0, -6).Copy
                Sheets("A" & i).Select
                Range("K3").PasteSpecial xlPasteAll
                
                Range("h3:h6").Replace What:="y", Replacement:="Y"
                Range("L3").Formula2R1C1 = "=IFERROR(ROUND((R3C10/R3C9)*(1+R3C17),3),"""")"
                Range("J3").Formula2R1C1 = "=IFERROR(ROUND(R4C10+((R5C10+(R6C10/32))/36),3),"""")"
                Range("J4").FormulaR1C1 = "=IFERROR(LEFT(R[-1]C[-2],FIND(""Y"",R[-1]C[-2],1)-1),0)"
                Range("J5").Formula = "=IFERROR(MID(H3,FIND(""d"",H3,1)+1,LEN(H3)-FIND(""d"",H3,1)-3)+BUTTON!F3,""QWERTY"")"
                
                Range("J6").FormulaR1C1 = "=IFERROR(RIGHT(R[-3]C[-2],2),0)"
                Range("J4:J6").Select
                Selection.Copy
                Selection.PasteSpecial xlPasteValues
                Application.CutCopyMode = False
                
                Columns("A:O").AutoFit
                
                If Cells(5, 10) = "QWERTY" Then
                    
'                    Cells(10, 1) = Cells(3, 8)
'
'                    Cells(10, 1).TextToColumns Destination:=Range("A10"), DataType:=xlDelimited, _
'                    TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
'                    Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
'                    :=",", FieldInfo:=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
'
'                    Cells(11, 1) = Cells(10, 1)
'                    Cells(11, 1).TextToColumns Destination:=Range("A11"), DataType:=xlDelimited, _
'                    TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
'                    Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
'                    :="d", FieldInfo:=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
'
'                    Cells(6, 10) = Cells(10, 2)
'                    Cells(5, 10) = Cells(11, 2)
'
'                    Range("A10:T20").ClearContents

                    'NEW SAM 30-JULI-2024
                    Cells(15, 1) = Cells(3, 8)
                    Set rng = Cells(15, 1)
    
                    ' Daftar delimiter yang umum digunakan
                    delimiters = "= Yd,.;:|/\~!@#$%^&*()_-+{}[]<>?`'"""
                    
                    ' Loop untuk mengganti semua delimiter dengan satu pembatas '|'
                    For k = 1 To Len(delimiters)
                        currentChar = Mid(delimiters, k, 1)
                        rng.Value = Replace(rng.Value, currentChar, "|")
                    Next k
                    
                    Data = rng.Value
                    delimiter = "|"
    
                    ' Ganti semua urutan delimiter berturut-turut dengan satu delimiter
                    Do While InStr(Data, delimiter & delimiter) > 0
                        Data = Replace(Data, delimiter & delimiter, delimiter)
                    Loop
                    
                    Cells(15, 1) = Data
                    
                    ' Pisahkan teks berdasarkan pembatas '|'
                    rng.TextToColumns Destination:=rng, _
                                      DataType:=xlDelimited, _
                                      TextQualifier:=xlDoubleQuote, _
                                      ConsecutiveDelimiter:=True, _
                                      Other:=True, _
                                      OtherChar:="|", _
                                      FieldInfo:=Array(1, 1), _
                                      TrailingMinusNumbers:=True

                    Cells(4, 10) = Cells(15, 1)
                    Cells(5, 10) = Cells(15, 2)
                    Cells(6, 10) = Cells(15, 3)
                    
                    Rows(15).ClearContents
                    Set rng = Nothing
                End If
                
            End If
        Next j
    End If
Next i

For i = 1 To SUMNO
    Sheets("TES").Select
    If Cells(i, 2) = "" Then
        For j = 30 To SUMALL
            A = Sheets("TES").Cells(i, 1).Value
            B = Sheets("MA use").Cells(j, 17).Value
            If A = B Then
                Sheets("A" & i).Select
                Range("J3").Copy
                Sheets("MA").Select
                Cells(j, 10).Select
                ActiveSheet.Paste Link:=True
                Sheets("A" & i).Select
                Range("L3").Copy
                Sheets("MA").Select
                Cells(j, 12).Select
                ActiveSheet.Paste Link:=True
            End If
        Next j
    End If
Next i

Sheets("MA use").Select
Range("I" & Rows.Count).End(xlUp).Select
Range(Selection, "I30").Copy
Sheets("MA").Select
Range("I30").PasteSpecial xlPasteAll

'Sheets("MA use").Select
'Range("J" & Rows.Count).End(xlUp).Select
'Range(Selection, "J30").Copy
'Sheets("MA").Select
'Range("J30").PasteSpecial xlPasteAll

'Sheets("MA use").Select
'Range("L" & Rows.Count).End(xlUp).Select
'Range(Selection, "L30").Copy
'Sheets("MA").Select
'Range("L30").PasteSpecial xlPasteAll

Sheets("FORMAT").Select
ActiveWindow.SelectedSheets.Visible = False

Application.DisplayAlerts = False
Sheets("MA use").Delete
Sheets("TES").Delete
Sheets("TES1").Delete
Application.DisplayAlerts = True

Sheets("MA").Select
Range("A30").Select

'ActiveWorkbook.Save

elapsedtime = Round(Timer - starttime, 2)
MsgBox "Successfully done in " & elapsedtime & " seconds", , "PERHITUNGAN CONS"


End Sub

