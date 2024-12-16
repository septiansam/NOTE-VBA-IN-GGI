
'{ DEVELOPER:= SEPTIAN ARIF MAULANA }
'_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-

Dim TWB As Workbook, HOME As Worksheet, INPUT_USER As Worksheet
Dim TEMP As Worksheet, UPLOAD As Worksheet, RPA As Worksheet

Dim WB_DB As Workbook, SH_DB As Worksheet
Dim RNG_DATA As Range, RNG_UPLOAD As Range
Dim PATH_DB As String
Dim STR_FILE_NAME As String
Dim STR_SH_NAME As String
Dim SUM_DATA As Long

Dim WB_CSV As Workbook
Dim FILE_CSV As String
Dim PATH_CSV As String
Dim PATH_UPLOAD As String

Dim WB_RESUME As Workbook, RNG_RESUME As Range
Dim PATH_RESUME As String

Dim WB_HISTORY As Workbook
Dim STR_HISTORY_RESUME As String
Dim PATH_HISTORY As String

Dim i As Long, j As Long, LR As Long, LC As Long

Dim isFound As Boolean '# UNTUK MENGECEK WORKBOOK
Dim isData As Boolean  '# UNTUK MENGECEK WORKSHEETS

Dim A As Long, B As Long, C As Long
Dim STR_STYLE_BARU As String
Dim STR_SPECIAL As String
Dim DATE_DELIVERY As Date
Dim STR_PO As String
Dim QTY_ORDER As Long
Dim UNIT_PRC As Double
Dim STYLE_BASE As String
Dim STR_KATEGORI As String
Dim STR_COLOR As String

Dim SUM_KEDATANGAN As Long, SUM_INSERT As Long
Dim STR_ITEM_KEDATANGAN As String
Dim RNG_ITEM As Range
Dim RNG_DETAIL As Range
Dim COL_PASTE As Long

Dim ANSWER As VbMsgBoxResult

Dim STR_UNIQ_KEY As String
Dim STR_STYLES As String, COUNT_STYLE As Long
Dim CEK_STATUS As String

Dim CELL As Range, RNG_COL As Range, COLUMN_COUNT As Long

Sub PROSES1()

Application.DisplayAlerts = False

Set TWB = ThisWorkbook
Set HOME = TWB.Sheets("HOME")
Set INPUT_USER = TWB.Sheets("INPUT USER")

'[VALIDASI]
'``````````````````````````````````````````````````````````````````````````
With INPUT_USER
    A = Application.WorksheetFunction.CountA(.Range("A:A")) - 1
    B = Application.WorksheetFunction.CountA(.Range("B:B")) - 1
    C = Application.WorksheetFunction.CountA(.Range("C:C")) - 1
End With
If A < B And A < C Then
    HOME.Activate: Cells(1, 1).Select
    MsgBox "Terdapat STYLE BARU Yang Kosong", vbExclamation, "Not Found..."
    Exit Sub
ElseIf B < A And B < C Then
    HOME.Activate: Cells(1, 1).Select
    MsgBox "Terdapat STYLE BASE Yang Kosong", vbExclamation, "Not Found..."
    Exit Sub
ElseIf C < A And C < B Then
    HOME.Activate: Cells(1, 1).Select
    MsgBox "Terdapat KATEGORI Yang Kosong", vbExclamation, "Not Found..."
    Exit Sub
End If

'[END VALIDASI]
'__________________________________________________________________________

For i = Sheets.Count To 3 Step -1
    Sheets(i).Delete
Next i

If WorksheetExists("RPA") Then Sheets("RPA").Delete
If WorksheetExists("TEMP") Then Sheets("TEMP").Delete

Set RPA = Sheets.Add(After:=Sheets(Sheets.Count)): ActiveSheet.Name = "RPA"
Set TEMP = Sheets.Add(After:=Sheets(Sheets.Count)): ActiveSheet.Name = "TEMP"

RPA.Activate
Rows(1).Font.Bold = True

Range("A1").Value = "STYLE"
Range("B1").Value = "SPECIAL"
Range("C1").Value = "PARENT ITEM"
Range("D1").Value = "2ND ITEM NUMBER"
Range("E1").Value = "DELIVERY"
Range("F1").Value = "PO"
Range("G1").Value = "QTY ORDER"
Range("H1").Value = "UNIT PRICE"
Range("I1").Value = "OR"
Range("J1").Value = "UPLOAD"
Range("K1").Value = "STYLE BASE"
Range("L1").Value = "KATEGORI"
Range("M1").Value = "COLOR"

'[KEDATANGAN]........................................
'```````````````````````````````````````````````````
INPUT_USER.Activate
LR = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
LC = Cells(1, Columns.Count).End(xlToLeft).Column

Set RNG_ITEM = Range(Cells(3, 10), Cells(LR, LC))

Range("J1", Cells(1, LC)).Copy

RPA.Activate
COL_PASTE = Cells(1, Columns.Count).End(xlToLeft).Column + 1
Cells(1, COL_PASTE).PasteSpecial xlPasteValuesAndNumberFormats

RNG_ITEM.Copy
Cells(2, COL_PASTE).PasteSpecial xlPasteValuesAndNumberFormats

Application.CutCopyMode = False
Cells.WrapText = False
Cells.HorizontalAlignment = xlCenter
Cells.VerticalAlignment = xlCenter
Cells.EntireColumn.AutoFit
Cells(1, 1).Select

' HAPUS SEMUA FILE CSV YANG ADA DI DIREKTORI FILE CSV
Dim file As String

PATH_CSV = HOME.Range("E6")

FILE_CSV = Dir(PATH_CSV & Application.PathSeparator & "UPLOAD" & "*.*")

Do While FILE_CSV <> ""
    ' Menggunakan Kill untuk menghapus file
    Kill PATH_CSV & Application.PathSeparator & FILE_CSV
    
    ' Mengambil file berikutnya dalam direktori
    FILE_CSV = Dir
Loop

SUM_DATA = Application.WorksheetFunction.CountA(INPUT_USER.Range("B3:B" & Rows.Count))

For i = 1 To SUM_DATA
    TWB.Activate
    TEMP.Cells.ClearContents
    
    If WorksheetExists("UPLOAD_" & i) Then Sheets("UPLOAD_" & i).Delete
    
    STR_FILE_NAME = Left(INPUT_USER.Range("B" & i + 2), 6) & _
                    INPUT_USER.Range("C" & i + 2)
    PATH_DB = HOME.Range("E5") & Application.PathSeparator & STR_FILE_NAME & ".xlsx"
    STR_SH_NAME = INPUT_USER.Range("B" & i + 2)
    
    If Dir(PATH_DB) = "" Then
        Stop
    End If
    
    Set WB_DB = Workbooks.Open(PATH_DB)
    
    WB_DB.Activate
    If Not WorksheetExists(STR_SH_NAME) Then
        Stop
    End If
    
    Set SH_DB = WB_DB.Sheets(STR_SH_NAME)
    SH_DB.Activate
    LR = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    LC = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    Set RNG_DATA = Range(Cells(1, 1), Cells(LR, LC))
    
    TEMP.Activate
    RNG_DATA.Copy
    
    Range("A1").PasteSpecial xlPasteAll: Application.CutCopyMode = False: Cells.EntireColumn.AutoFit: Cells(1, 1).Select
    WB_DB.Close (False)
    
    '[CLEAN DATA, GET RANGE DETAIL]
    '`````````````````````````````````````````````````````````````````````````````````````````````````````````````````````````````
    TEMP.Activate
    Range("A:B").Delete: Cells(1, 1).Select
    LC = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column

    Set RNG_DETAIL = Range(Cells(2, 2), Cells(2, LC))
    
    '[APAKAH APAKAH ADA KEDATANGAN ITEM BARU]
    '`````````````````````````````````````````````````````````````````````````````````````````````````````````````````````````````
    INPUT_USER.Activate
    SUM_KEDATANGAN = Application.WorksheetFunction.CountA(Range(Cells(i + 2, 10), Cells(i + 2, Columns.Count)))
    
    If SUM_KEDATANGAN > 0 Then
        SUM_INSERT = SUM_KEDATANGAN - 1
        TEMP.Activate
        If SUM_INSERT > 0 Then
            Rows("2:" & 1 + SUM_INSERT).Insert
        End If
        
        For j = 1 To SUM_KEDATANGAN
            STR_ITEM_KEDATANGAN = CStr(INPUT_USER.Cells(i + 2, 9 + j).Value)
            Cells(j + 1, 1).Value = STR_ITEM_KEDATANGAN
            RNG_DETAIL.Copy
            Cells(j + 1, 2).PasteSpecial xlPasteAll: Application.CutCopyMode = False: Cells.EntireColumn.AutoFit: Cells(1, 1).Select
        Next j
    End If
    TEMP.Activate
    Cells.EntireColumn.AutoFit: Cells(1, 1).Select

    LR = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    LC = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    Set RNG_UPLOAD = Range(Cells(2, 1), Cells(LR, LC))
    
    Set UPLOAD = TWB.Sheets.Add(After:=TWB.Sheets(TWB.Sheets.Count))
    ActiveSheet.Name = "UPLOAD_" & i
    
    UPLOAD.Activate
    RNG_UPLOAD.Copy
    Range("A1").PasteSpecial xlPasteAll: Application.CutCopyMode = False: Cells.EntireColumn.AutoFit: Cells(1, 1).Select
    
    '[SAVE UPLOAD CSV]...
    '````````````````````
    FILE_CSV = "UPLOAD_" & i
    UPLOAD.Copy
    Set WB_CSV = ActiveWorkbook
    WB_CSV.Activate
    WB_CSV.SaveAs fileName:=PATH_CSV & Application.PathSeparator & FILE_CSV, _
                    FileFormat:=xlCSV
    WB_CSV.Close SaveChanges:=True
    
    '[PUT DATA INTO SHEETS RPA]
    '````````````````````
    INPUT_USER.Activate
    
    Cells(1, 1).Select
    With INPUT_USER
        STR_STYLE_BARU = CStr(.Range("A" & i + 2).Value)
        STR_SPECIAL = CStr(.Range("D" & i + 2).Value)
        DATE_DELIVERY = Format(CDate(.Range("F" & i + 2).Value), "m/d/yyyy")
        STR_PO = CStr(.Range("G" & i + 2).Value)
        QTY_ORDER = CLng(.Range("H" & i + 2).Value)
        UNIT_PRC = CDbl(.Range("I" & i + 2).Value)
        STYLE_BASE = CStr(.Range("B" & i + 2).Value)
        STR_KATEGORI = CStr(.Range("C" & i + 2).Value)
        STR_COLOR = CStr(.Range("E" & i + 2).Value)
    End With
    PATH_UPLOAD = PATH_CSV & Application.PathSeparator & FILE_CSV & ".csv"

    RPA.Activate

    Range("A" & i + 1).Value = STR_STYLE_BARU
    Range("B" & i + 1).Value = STR_SPECIAL
    Range("E" & i + 1).Value = DATE_DELIVERY: Range("E" & i + 1).NumberFormat = "m/d/yyyy"
    Range("F" & i + 1).Value = STR_PO
    Range("G" & i + 1).Value = QTY_ORDER
    Range("H" & i + 1).Value = UNIT_PRC: Range("H" & i + 1).NumberFormat = "_($* #,##0.00_);[Red]_($* (#,##0.00);_($* ""-""??_);_(@_)"

    RPA.Hyperlinks.Add _
                    Anchor:=Range("J" & i + 1), _
                    Address:=PATH_UPLOAD, _
                    TextToDisplay:=PATH_UPLOAD
    
    Range("K" & i + 1).FormulaR1C1 = STYLE_BASE
    Range("L" & i + 1).FormulaR1C1 = STR_KATEGORI
    Range("M" & i + 1).FormulaR1C1 = STR_COLOR
    
    Cells.EntireColumn.AutoFit
    
    Set UPLOAD = Nothing
    Set WB_CSV = Nothing
    
    HOME.Activate
Next i

'[END].................................................................
'`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.

TEMP.Delete
HOME.Activate: Cells(1, 1).Select

Application.DisplayAlerts = True

End Sub

Sub PROSES2()

Application.DisplayAlerts = False

Set TWB = ThisWorkbook
Set HOME = TWB.Sheets("HOME")
Set INPUT_USER = TWB.Sheets("INPUT USER")

'[VALIDASI]______________________________________________________
'````````````````````````````````````````````````````````````````
If Not WorksheetExists("RPA") Then
    HOME.Activate: Cells(1, 1).Select
    MsgBox "PROSES 1 BELUM DIJALANKAN", vbExclamation, "PROCESS 1 HAS NOT BEEN EXECUTED"
    Exit Sub
End If
Set RPA = TWB.Sheets("RPA")
RPA.Activate
If Application.WorksheetFunction.CountA(Rows("2:" & Rows.Count)) = 0 Then
    HOME.Activate: Cells(1, 1).Select
    MsgBox "THERE IS NO DATA ON THE RPA SHEET", vbExclamation, "NOT FOUND"
    Exit Sub
End If
With RPA
    A = Application.WorksheetFunction.CountA(.Range("C:C")) - 1
    B = Application.WorksheetFunction.CountA(.Range("D:D")) - 1
    C = Application.WorksheetFunction.CountA(.Range("I:I")) - 1
End With
If A + B + C = 0 Then
    HOME.Activate: Cells(1, 1).Select
    MsgBox "DATA PARENT ITEM, ITEM NUMBER, OR IS EMPTY", vbExclamation, "Not Found..."
    Exit Sub
End If
If A < B And A < C Then
    HOME.Activate: Cells(1, 1).Select
    MsgBox "Terdapat PARENT ITEM Yang Kosong", vbExclamation, "Not Found..."
    Exit Sub
ElseIf B < A And B < C Then
    HOME.Activate: Cells(1, 1).Select
    MsgBox "Terdapat ITEM NUMBER Yang Kosong", vbExclamation, "Not Found..."
    Exit Sub
ElseIf C < A And C < B Then
    HOME.Activate: Cells(1, 1).Select
    MsgBox "Terdapat OR Yang Kosong", vbExclamation, "Not Found..."
    Exit Sub
End If
'[END VALIDASI]
'__________________________________________________________________________

'[SAVE RESUME FOR REPORT]...
'````````````````````````````
RPA.Activate

Cells.Interior.Color = xlNone
Cells.Borders.LineStyle = xlNone

LR = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
LC = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column

PATH_RESUME = HOME.Range("E7") & Application.PathSeparator & _
                HOME.Range("D7") & ".xlsx"

COLUMN_COUNT = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
For i = COLUMN_COUNT To 1 Step -1
    If Application.WorksheetFunction.CountA(Columns(i)) = 1 Then
        Columns(i).Delete
    End If
Next i

LC = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
Set RNG_DATA = Range(Cells(1, 1), Cells(LR, LC))

'[DESIGN].......
'```````````````
Call DISPLAY_DESIGN(LR, LC, RNG_DATA)
'```````````````

RPA.Copy
Set WB_RESUME = ActiveWorkbook
WB_RESUME.Activate: Sheets(1).Select
Columns(10).Delete: Cells.EntireColumn.AutoFit: Cells(1, 1).Select
WB_RESUME.SaveAs PATH_RESUME, xlOpenXMLWorkbook
WB_RESUME.Close (True)

'[SAVE RESUME FOR HISTORY]...
'````````````````````````````
RPA.Activate
COUNT_STYLE = Application.WorksheetFunction.CountA(Range("A:A")) - 1
STR_STYLES = ""
For i = 1 To COUNT_STYLE
    If STR_STYLES = "" Then
        STR_STYLES = Cells(i + 1, 1).Value
    Else
        STR_STYLES = STR_STYLES & ", " & Cells(i + 1, 1).Value
    End If
Next i
STR_UNIQ_KEY = STR_STYLES & "_" & Format(Now(), "MMDDYYYY_HHMMSS")
STR_HISTORY_RESUME = "RESUME" & "_" & STR_UNIQ_KEY
PATH_HISTORY = HOME.Range("AF2") & Application.PathSeparator & STR_HISTORY_RESUME & ".xlsx"

RPA.Copy
Set WB_HISTORY = ActiveWorkbook
WB_HISTORY.Activate: Sheets(1).Select

Cells.EntireColumn.AutoFit: Cells(1, 1).Select
WB_HISTORY.SaveAs PATH_HISTORY, xlOpenXMLWorkbook
WB_HISTORY.Close (True)

'[END]_______________________
'````````````````````````````
TWB.Activate

INPUT_USER.Activate

Cells.EntireColumn.AutoFit: Cells(1, 1).Select
HOME.Activate: Cells(1, 1).Select

End Sub

Sub CLEAR()
Application.DisplayAlerts = False

Set TWB = ThisWorkbook
Set HOME = TWB.Sheets("HOME")
Set INPUT_USER = TWB.Sheets("INPUT USER")

'ANSWER = MsgBox("Apakah Anda ingin menghapus?" _
'        & vbCrLf & vbCrLf & _
'        "Jika dihapus, inputan user akan hilang" _
'        & vbCrLf & vbCrLf & _
'        "Pilih Yes untuk melanjutkan atau No untuk keluar.", vbYesNo + vbQuestion, "*) Konfirmasi..................")
'
'If ANSWER = vbNo Then
'    Exit Sub
'End If

For i = TWB.Sheets.Count To 3 Step -1
    Sheets(i).Delete
Next i

INPUT_USER.Activate
Rows("3:100").ClearContents
Cells(1, 1).Select

HOME.Activate: Cells(1, 1).Select
HOME.Range("AD2").Value = ""

Application.DisplayAlerts = True
End Sub

Sub DISPLAY_DESIGN(ByRef LastRow, ByRef LastCol, Range_Data As Range)

    With Range(Cells(1, 1), Cells(1, LC))
        .HorizontalAlignment = xlCenter
        .Font.Name = "Century Gothic"
        .Font.Color = vbBlack
        .Font.Bold = True
        .Font.Size = 13
        .Interior.Color = RGB(28, 136, 226)
        .Interior.Pattern = xlSolid
        .Interior.PatternColorIndex = xlAutomatic
        .RowHeight = .RowHeight + 10
    End With
    
    Cells.EntireColumn.AutoFit: Cells(1, 1).Select
    For i = 2 To LR
        If i Mod 2 = 0 Then
            With Range(Cells(i, 1), Cells(i, LC))
                .Interior.Pattern = xlSolid
                .Interior.PatternColor = xlAutomatic
                .Interior.Color = RGB(255, 255, 255)
                .RowHeight = .RowHeight + 2
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeTop).Color = RGB(5, 160, 255)
                .Borders(xlEdgeTop).TintAndShade = 0
                .Borders(xlEdgeTop).Weight = xlThin
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).Color = RGB(5, 160, 255)
                .Borders(xlEdgeBottom).TintAndShade = 0
                .Borders(xlEdgeBottom).TintAndShade = 0
            End With
        ElseIf i Mod 2 <> 0 Then
            With Range(Cells(i, 1), Cells(i, LC))
                .Interior.Pattern = xlSolid
                .Interior.PatternColor = xlAutomatic
                .Interior.Color = RGB(228, 240, 241)
                .Interior.Color = RGB(214, 234, 250)
                .RowHeight = .RowHeight + 2
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeTop).Color = RGB(5, 160, 255)
                .Borders(xlEdgeTop).TintAndShade = 0
                .Borders(xlEdgeTop).Weight = xlThin
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).Color = RGB(5, 160, 255)
                .Borders(xlEdgeBottom).TintAndShade = 0
                .Borders(xlEdgeBottom).TintAndShade = 0
            End With
        End If
    Next i
    
    For Each RNG_COL In Range_Data.Columns
        With RNG_COL
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeLeft).Color = vbWhite
            .Borders(xlEdgeLeft).TintAndShade = 0
            .Borders(xlEdgeLeft).TintAndShade = 0
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlEdgeRight).Color = vbWhite
            .Borders(xlEdgeRight).TintAndShade = 0
            .Borders(xlEdgeRight).TintAndShade = 0
            .Borders(xlInsideVertical).LineStyle = xlContinuous
            .Borders(xlInsideVertical).Color = vbWhite
            .Borders(xlInsideVertical).TintAndShade = 0
            .Borders(xlInsideVertical).TintAndShade = 0
        End With
        RNG_COL.ColumnWidth = Range_Data.ColumnWidth + 1
    Next RNG_COL
    
    Cells(1, 1).Select

End Sub


''''[ FUNGSI CEK SHEET ]''''
Function WorksheetExists(shtName As String) As Boolean
    On Error Resume Next
        WorksheetExists = Not Sheets(shtName) Is Nothing
    On Error GoTo 0
End Function

