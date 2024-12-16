Sub PERSEN_LUR()
Dim TWB As Workbook, WS1 As Worksheet, i As Integer, j As Integer
Dim TES1 As String, TES2 As String, TES3 As String

Set TWB = ThisWorkbook: Set WS1 = TWB.Sheets("TOMBOL")
TES1 = "TES1": TES2 = "TES2": TES3 = "TES3"

Call APUSS

For i = 1 To 3
Sheets.Add(After:=Sheets(Sheets.Count)).Name = "TES" & i
Next i

FILEPATHWO = WS1.Range("F4").Value & Application.PathSeparator & WS1.Range("F5").Value & ".xlsx": FILEWO = WS1.Range("F5").Value & ".xlsx"
DirFile = FILEPATHWO
If Dir(DirFile) = "" Then
    TWB.Activate
    Call APUSS
    MsgBox "File " & FILEWO & " doesn't exist", vbCritical: Exit Sub
Else
    Application.DisplayAlerts = False: Workbooks.Open FILEPATHWO
End If
Sheets(1).Select

Dim rgSource As Range
Set rgSource = Range("A1").CurrentRegion
rgSource.Copy: TWB.Sheets("TES1").Cells(1, 1).PasteSpecial xlPasteValues: Application.CutCopyMode = False

Workbooks(FILEWO).Close SAVECHANGES:=False
Sheets(TES1).Select

Columns("N").Copy Destination:=Sheets(TES2).Cells(1, 1)
Sheets(TES2).Select
Range("A1:A100000").RemoveDuplicates Columns:=1, Header:=xlYes

For i = Range("A" & Rows.Count).End(xlUp).Row To 2 Step -1
    If Cells(i, 1) = vbNullString Then
        Rows(i).Delete
    End If
Next i

LAST_TES2 = Sheets("TES2").Cells(Rows.Count, 1).End(xlUp).Row
'WORKSHEET
Range("B2:B" & LAST_TES2).Formula = "=COUNTIF('TES1'!N:N,'TES2'!A2)"
Range("C2:C" & LAST_TES2).Formula = "=COUNTIFS('TES1'!N:N,'TES2'!A2,'TES1'!O:O,""1"")"
Range("D2:D" & LAST_TES2).Formula = "=(C2/B2)*100%"
Range("E2:E" & LAST_TES2).Formula = "=COUNTIFS('TES1'!N:N,'TES2'!A2,'TES1'!O:O,""0"")"
Range("F2:F" & LAST_TES2).Formula = "=(E2/B2)*100%"

'TRIMCARD
Range("G2:G" & LAST_TES2).Formula = "=COUNTIF('TES1'!Q:Q,'TES2'!A2)"
Range("H2:H" & LAST_TES2).Formula = "=COUNTIFS('TES1'!Q:Q,'TES2'!A2,'TES1'!R:R,""1"")"
Range("I2:I" & LAST_TES2).Formula = "=(H2/G2)*100%"
Range("J2:J" & LAST_TES2).Formula = "=COUNTIFS('TES1'!Q:Q,'TES2'!A2,'TES1'!R:R,""0"")"
Range("K2:K" & LAST_TES2).Formula = "=(J2/G2)*100%"

'SAMPLE
Range("L2:L" & LAST_TES2).Formula = "=COUNTIF('TES1'!T:T,'TES2'!A2)"
Range("M2:M" & LAST_TES2).Formula = "=COUNTIFS('TES1'!T:T,'TES2'!A2,'TES1'!U:U,""1"")"
Range("N2:N" & LAST_TES2).Formula = "=(M2/L2)*100%"
Range("O2:O" & LAST_TES2).Formula = "=COUNTIFS('TES1'!T:T,'TES2'!A2,'TES1'!U:U,""0"")"
Range("P2:P" & LAST_TES2).Formula = "=(O2/L2)*100%"


Range("B1") = "Total Worksheet": Range("C1") = "Worksheet Release": Range("E1") = "Worksheet not Release"
Range("G1") = "Total Trimcard": Range("H1") = "Trimcard Release": Range("J1") = "Trimcard not Release"
Range("L1") = "Total Sample": Range("M1") = "Sample Release": Range("O1") = "Sample not Release"

For i = 16 To 4 Step -1
If Cells(1, i) = vbNullString Then
    Range(Cells(1, i), Cells(1, i - 1)).Merge
    Range(Cells(1, i), Cells(1, i - 1)).HorizontalAlignment = xlCenter
    
    Range(Cells(2, i), Cells(LAST_TES2, i)).NumberFormat = "0%"
End If
Next i
Range("A1").CurrentRegion.HorizontalAlignment = xlCenter
Set Rng = Range("A1").CurrentRegion
With Rng.Borders 'BUAT BORDER
    .LineStyle = xlContinuous
    .Weight = xlThin
End With
'===================================
Sheets("TES1").Select
Columns("W").Copy Destination:=Sheets(TES3).Cells(1, 1)
Sheets(TES3).Select
Range("A1:A100000").RemoveDuplicates Columns:=1, Header:=xlYes
For i = Range("A" & Rows.Count).End(xlUp).Row To 2 Step -1
    If Cells(i, 1) = vbNullString Then
        Rows(i).Delete
    End If
Next i
DATA_TES3 = Cells(Rows.Count, 1).End(xlUp).Row
Range("B2:B" & DATA_TES3).Formula = "=COUNTIF('TES1'!W:W,'TES3'!A2)"
Range("C2:C" & DATA_TES3).Formula = "=COUNTIFS('TES1'!W:W,'TES3'!A2,'TES1'!X:X,""1"")"
Range("D2:D" & DATA_TES3).Formula = "=(C2/B2)*100%"
Range("E2:E" & DATA_TES3).Formula = "=COUNTIFS('TES1'!W:W,'TES3'!A2,'TES1'!X:X,""0"")"
Range("F2:F" & DATA_TES3).Formula = "=(E2/B2)*100%"

Range("B1") = "Total PilotRun": Range("C1") = "PilotRun Release": Range("E1") = "PilotRun not Release"

For i = 6 To 4 Step -1
If Cells(1, i) = vbNullString Then
    Range(Cells(1, i), Cells(1, i - 1)).Merge
    Range(Cells(1, i), Cells(1, i - 1)).HorizontalAlignment = xlCenter
    
    Range(Cells(2, i), Cells(LAST_TES2, i)).NumberFormat = "0%"
End If
Next i
Range("A1").CurrentRegion.HorizontalAlignment = xlCenter
Set Rng = Range("A1").CurrentRegion
With Rng.Borders 'BUAT BORDER
    .LineStyle = xlContinuous
    .Weight = xlThin
End With
Cells.Copy: Cells(1, 1).PasteSpecial xlPasteValues: Application.CutCopyMode = False
Rng.Copy Destination:=Sheets(TES2).Range("A" & Rows.Count).End(xlUp).Offset(2, 0)
Cells.Delete
'=================================================
Sheets("TES1").Select
Columns("Z").Copy Destination:=Sheets(TES3).Cells(1, 1)
Sheets(TES3).Select
Range("A1:A100000").RemoveDuplicates Columns:=1, Header:=xlYes
For i = Range("A" & Rows.Count).End(xlUp).Row To 2 Step -1
    If Cells(i, 1) = vbNullString Then
        Rows(i).Delete
    End If
Next i

Range("A1").CurrentRegion.Copy Destination:=Sheets(TES2).Range("A" & Rows.Count).End(xlUp).Offset(2, 0)

Sheets(TES2).Select
DATATES2 = Range("A" & Rows.Count).End(xlUp).Row
BARISPOLA = Cells.Find("PolaMarker.PIC").Row

Cells(BARISPOLA, 2) = "Total PolaMarker": Cells(BARISPOLA, 3) = "PolaMarker Release": Cells(BARISPOLA, 5) = "PolaMarker not Release"
Range(Cells(BARISPOLA + 1, 2), Cells(DATATES2, 2)).Formula = "=COUNTIF('TES1'!Z:Z,'TES2'!A" & BARISPOLA + 1 & ")"
Range(Cells(BARISPOLA + 1, 3), Cells(DATATES2, 3)).Formula = "=COUNTIFS('TES1'!Z:Z,'TES2'!A" & BARISPOLA + 1 & ",'TES1'!AA:AA,""1"")"
Range(Cells(BARISPOLA + 1, 4), Cells(DATATES2, 4)).Formula = "=(C" & BARISPOLA + 1 & "/B" & BARISPOLA + 1 & ")*100%"
Range(Cells(BARISPOLA + 1, 5), Cells(DATATES2, 5)).Formula = "=COUNTIFS('TES1'!Z:Z,'TES2'!A" & BARISPOLA + 1 & ",'TES1'!AA:AA,""0"")"
Range(Cells(BARISPOLA + 1, 6), Cells(DATATES2, 6)).Formula = "=(E" & BARISPOLA + 1 & "/B" & BARISPOLA + 1 & ")*100%"

For i = 6 To 4 Step -1
If Cells(1, i) = vbNullString Then
    Range(Cells(BARISPOLA, i), Cells(BARISPOLA, i - 1)).Merge
    Range(Cells(BARISPOLA, i), Cells(BARISPOLA, i - 1)).HorizontalAlignment = xlCenter
    
    Range(Cells(BARISPOLA + 1, i), Cells(DATATES2, i)).NumberFormat = "0%"
End If
Next i
Set Rng = Range("A" & BARISPOLA + 1).CurrentRegion
With Rng.Borders 'BUAT BORDER
    .LineStyle = xlContinuous
    .Weight = xlThin
End With
'=====================================
Sheets("TES3").Cells.Delete
Sheets("TES1").Select
Columns("AC").Copy Destination:=Sheets(TES3).Cells(1, 1)
Sheets(TES3).Select
Range("A1:A100000").RemoveDuplicates Columns:=1, Header:=xlYes
For i = Range("A" & Rows.Count).End(xlUp).Row To 2 Step -1
    If Cells(i, 1) = vbNullString Then
        Rows(i).Delete
    End If
Next i
Range("A1").CurrentRegion.Copy Destination:=Sheets(TES2).Range("A" & Rows.Count).End(xlUp).Offset(2, 0)

Sheets(TES2).Select
DATATES2 = Range("A" & Rows.Count).End(xlUp).Row
BARISKONKER = Cells.Find("Konker.PIC").Row
Cells(BARISKONKER, 2) = "Total Konker": Cells(BARISKONKER, 3) = "Konker Release": Cells(BARISKONKER, 5) = "Konker not Release"
'Range(Cells(BARISKONKER + 1, 2), Cells(DATATES2, 2)).Formula = "=COUNTIF('TES1'!AC:AC,'TES2'!A" & BARISKONKER + 1 & ")"
Range(Cells(BARISKONKER + 1, 2), Cells(DATATES2, 2)).FormulaR1C1 = _
    "=IF(RC[-1]=""Fauzi"",COUNTIFS('TES1'!C[27],""Fauzi"",'TES1'!C[6],""<>MJ2"",'TES1'!C[12],""<>""),COUNTIF('TES1'!C[27],'TES2'!RC[-1]))"

'Range(Cells(BARISKONKER + 1, 3), Cells(DATATES2, 3)).Formula = "=COUNTIFS('TES1'!AC:AC,'TES2'!A" & BARISKONKER + 1 & ",'TES1'!AD:AD,""1"")"
Range(Cells(BARISKONKER + 1, 3), Cells(DATATES2, 3)).FormulaR1C1 = _
    "=IF(RC[-2]=""Fauzi"",COUNTIFS('TES1'!C[26],""Fauzi"",'TES1'!C[27],1,'TES1'!C[5],""<>MJ2"",'TES1'!C[11],""<>""),COUNTIFS('TES1'!C[26],'TES2'!RC[-2],'TES1'!C[27],""1""))"

Range(Cells(BARISKONKER + 1, 4), Cells(DATATES2, 4)).Formula = "=(C" & BARISKONKER + 1 & "/B" & BARISKONKER + 1 & ")*100%"
'Range(Cells(BARISKONKER + 1, 5), Cells(DATATES2, 5)).Formula = "=COUNTIFS('TES1'!AC:AC,'TES2'!A" & BARISKONKER + 1 & ",'TES1'!AD:AD,"""")"
Range(Cells(BARISKONKER + 1, 5), Cells(DATATES2, 5)).FormulaR1C1 = "=RC[-3]-RC[-2]"
Range(Cells(BARISKONKER + 1, 6), Cells(DATATES2, 6)).Formula = "=(E" & BARISKONKER + 1 & "/B" & BARISKONKER + 1 & ")*100%"

For i = 6 To 4 Step -1
If Cells(1, i) = vbNullString Then
    Range(Cells(BARISKONKER, i), Cells(BARISKONKER, i - 1)).Merge
    Range(Cells(BARISKONKER, i), Cells(BARISKONKER, i - 1)).HorizontalAlignment = xlCenter
    
    Range(Cells(BARISKONKER + 1, i), Cells(DATATES2, i)).NumberFormat = "0%"
End If
Next i

Set Rng = Range("A" & BARISKONKER + 1).CurrentRegion
With Rng.Borders 'BUAT BORDER
    .LineStyle = xlContinuous: .Weight = xlThin
End With

Columns("A:B").ColumnWidth = 15
Rows(1).Insert
Cells.Copy: Cells(1, 1).PasteSpecial xlPasteValues: Application.CutCopyMode = False
Cells(1, 1).Select

ActiveSheet.Name = "WO Preparation"
Call newfile
TES1 = "TES1": TES2 = "TES2": TES3 = "TES3"
Application.DisplayAlerts = False
If Evaluate("isref('" & TES1 & "'!A1)") Then
    Sheets(TES1).Delete
End If
If Evaluate("isref('" & TES2 & "'!A1)") Then
    Sheets(TES2).Delete
End If
If Evaluate("isref('" & TES3 & "'!A1)") Then
    Sheets(TES3).Delete
End If

WS1.Select

End Sub