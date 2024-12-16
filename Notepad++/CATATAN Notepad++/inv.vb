Sub invoice_agron()
Dim TWB As Workbook, WS1 As Worksheet, i As Integer, j As Integer, lastRow As Integer, DB2 As Worksheet

Set TWB = ThisWorkbook: Set WS1 = TWB.Sheets("INPUTAN"): Set DB1 = TWB.Sheets("DB2")

TES1 = "TES1": TES2 = "TES2"
PEKING = "INVOICE"

Application.DisplayAlerts = False
If Evaluate("isref('" & TES1 & "'!A1)") Then
    Sheets(TES1).Delete
End If
If Evaluate("isref('" & PEKING & "'!A1)") Then
    Sheets(PEKING).Delete
End If
If Evaluate("isref('" & TES2 & "'!A1)") Then
    Sheets(TES2).Delete
End If
TES3 = "TES3"
If Evaluate("isref('" & TES3 & "'!A1)") Then
    Application.DisplayAlerts = False
     Sheets(TES3).Delete
    Application.DisplayAlerts = True
End If
Sheets.Add(AFTER:=Sheets(Sheets.Count)).Name = "TES1": Sheets.Add(AFTER:=Sheets(Sheets.Count)).Name = "TES2"
Application.DisplayAlerts = True

Application.DisplayAlerts = False
LokasiFile = Application.GetOpenFilename(, , "AMBIL FILE LOADINGAN")
If LokasiFile = "False" Then
    MsgBox "TIDAK JADI", vbOKOnly + vbInformation, "GET Data": End
Else
    Workbooks.Open (LokasiFile)
End If
WB1 = ActiveWorkbook.Name

Workbooks(WB1).Activate
Sheets(1).Select

lastRow = Sheets(1).Cells.Find(What:="*" _
            , LookAt:=xlPart _
            , LookIn:=xlFormulas _
            , SearchOrder:=xlByRows _
            , SearchDirection:=xlPrevious).Row

Cells.Copy
TWB.Worksheets(TES1).Cells(1, 1).PasteSpecial xlPasteValues: Application.CutCopyMode = False

Workbooks(WB1).Close SAVECHANGES:=False
DB1.Visible = True
DB1.Select
Cells.Copy Destination:=Sheets(TES2).Cells(1, 1)

Sheets(TES1).Select 'PO DAN ITEM
DATATES1 = Cells(Rows.Count, 2).End(xlUp).Row

Range("A3:B" & DATATES1).Copy
Sheets(TES2).Select
Cells(28, 2).PasteSpecial xlPasteValues: Application.CutCopyMode = False

Sheets(TES1).Select 'KOLOR
Range("U3:U" & DATATES1).Copy
Sheets(TES2).Select
Cells(28, 4).PasteSpecial xlPasteValues: Application.CutCopyMode = False

Sheets(TES1).Select 'TOTAL PACK
Range("F3:F" & DATATES1).Copy
Sheets(TES2).Select
Cells(28, 8).PasteSpecial xlPasteValues: Application.CutCopyMode = False

Sheets(TES1).Select 'TOTAL CTN
Range("I3:I" & DATATES1).Copy
Sheets(TES2).Select
Cells(28, 9).PasteSpecial xlPasteValues: Application.CutCopyMode = False

Sheets(TES1).Select 'PRICE PACK
Range("E3:E" & DATATES1).Copy
Sheets(TES2).Select
Cells(28, 10).PasteSpecial xlPasteValues: Application.CutCopyMode = False

For i = 28 To Range("D" & Rows.Count).End(xlUp).Row
    If Right(Cells(i, 3), 1) = "A" Then
        Cells(i, 5) = "S"
    ElseIf Right(Cells(i, 3), 1) = "B" Then
        Cells(i, 5) = "M"
    ElseIf Right(Cells(i, 3), 1) = "C" Then
        Cells(i, 5) = "L"
    ElseIf Right(Cells(i, 3), 1) = "D" Then
        Cells(i, 5) = "XL"
    ElseIf Right(Cells(i, 3), 1) = "E" Then
        Cells(i, 5) = "XXL"
    End If
Next i '=H28*J28

Range("K28:K" & Range("D" & Rows.Count).End(xlUp).Row).Formula = "=H28*J28"

Sheets(TES1).Select 'desc
Range("AC3:AC" & DATATES1).Formula = "=AE3&"" ""&AD3"

Range("AC3:AC" & DATATES1).Copy
Sheets(TES2).Select
Cells(28, 6).PasteSpecial xlPasteValues: Application.CutCopyMode = False

For i = 501 To Range("D" & Rows.Count).End(xlUp).Row + 2 Step -1
    If Cells(i, 1) = vbNullString Then
        Rows(i).Delete
    End If
Next i

ActiveSheet.Name = "INVOICE"
Application.DisplayAlerts = False
If Evaluate("isref('" & TES1 & "'!A1)") Then
    Sheets(TES1).Delete
End If
Application.DisplayAlerts = True
Cells(1, 1).Select

DB1.Visible = False
Call Tambahan2

MsgBox "DONE", vbInformation


End Sub

Sub Tambahan2()

Dim WB_FSI As Workbook
Dim FSI As Worksheet, INV As Worksheet

Application.DisplayAlerts = False
Set TWB = ThisWorkbook
Set INPUTAN = TWB.Sheets("INPUTAN")
Set INV = TWB.Sheets("INVOICE")
Set PL = TWB.Sheets("Packing List")
Set WF = Application.WorksheetFunction

INV.Activate

INPUTAN.Activate
STR_FILE = INPUTAN.Range("I4")
PATH_FILE = INPUTAN.Range("I5") & Application.PathSeparator & STR_FILE & ".xls"
If Dir(PATH_FILE) = "" Then
    MsgBox "File " & STR_FILE & " Tidak Ditemukan", vbInformation, "File Not Found"
    End
    Exit Sub
End If
If WorksheetExists("FSI") Then Sheets("FSI").Delete
Set FSI = Sheets.Add(AFTER:=Sheets(Sheets.Count)): FSI.Name = "FSI"
Set WB_FSI = Workbooks.Open(PATH_FILE): WB_FSI.Activate: Sheets(1).Select
ActiveSheet.AutoFilterMode = False: Cells.EntireColumn.Hidden = False
Range("A:J").Copy
FSI.Activate
Range("A1").PasteSpecial xlPasteAll: Application.CutCopyMode = False
ActiveWindow.Zoom = 80
Cells.EntireColumn.AutoFit: Cells(1, 1).Select
WB_FSI.Close False

rDesc = Cells.Find("DESCRIPTION OF PACKAGES AND GOODS").Row
lr = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
lc = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column

Rows("" & rDesc & "" & ":" & "" & lr & "").Delete

val_Str = Range("E13").Value
If InStr(val_Str, "/") > 0 Then
    FSI.Range("E13").TextToColumns Destination:=FSI.Range("J13"), DataType:=xlDelimited, _
        Other:=True, OtherChar:="/"
    val_Flgt = Trim(Range("J13"))
    val_Trip = Trim(Range("K13"))
Else
    val_Flgt = ""
    val_Trip = ""
End If

'INVOICE#
val_Invoice = Range("H4")
val_DateInv = Range("H6")

'Orig Date
val_Date = Range("G13")
'val_DateOrigin = WF.Text(val_Date, "[$-id-ID]yyyy/mm/dd")

val_GW = PL.Range("N" & Rows.Count).End(xlUp).Value
val_NW = PL.Range("M" & Rows.Count).End(xlUp).Value

INV.Activate
Range("H4").Value = val_Invoice
Range("H5").Value = val_DateInv

Range("H17").Value = "Vsl/Flgt: " & UCase(val_Flgt)
Range("H18").Value = "Voy#/Trip: " & UCase(val_Trip)

Range("H20").Value = "Orig Dt: " & CStr(Format(val_Date, "yyyy/mm/dd"))

Range("K20").Value = CStr("GR WGT: " & val_GW)
Range("K21").Value = CStr("NET WGT: " & val_NW)

INV.Activate
lr = Range("A28").End(xlDown).Row
Set RNG = Range(Cells(28, 6), Cells(lr, 6))
For Each cell In RNG
    val_Str = cell.Value
    
    ' Check the conditions and update column B accordingly
    If InStr(1, val_Str, "cotton", vbTextCompare) > 0 Then
        If (InStr(1, val_Str, "crew neck", vbTextCompare) > 0 Or _
            InStr(1, val_Str, "tank top", vbTextCompare) > 0 Or _
            InStr(1, val_Str, "tshirt", vbTextCompare) > 0) Then
            cell.Offset(0, 1).Value = "61091010"
        ElseIf InStr(1, val_Str, "polyester", vbTextCompare) > 0 Then
            cell.Offset(0, 1).Value = "61071100"
        Else
            cell.Offset(0, 1).Value = "61071100"
        End If
    ElseIf InStr(1, val_Str, "polyester", vbTextCompare) > 0 Then
        cell.Offset(0, 1).Value = "61071200"
    End If
Next cell

ActiveWindow.Zoom = 100
Rows(9).RowHeight = 60
Cells.EntireColumn.AutoFit: Cells(1, 1).Select

FSI.Delete

INPUTAN.Activate: Cells(1, 1).Select

End Sub
