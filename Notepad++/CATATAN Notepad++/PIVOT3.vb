Sub CreatePivot3()
Dim twb As Workbook
Set twb = ThisWorkbook

Dim DataSrc As Worksheet
Dim SH_PIVOT As Worksheet
Dim PivotTable As PivotTable
Dim PivotCache As PivotCache
Dim PivotRange As Range
Dim lr As Long, lc As Long

Set DataSrc = twb.Sheets("BULAN") ' Ganti "DataSheet" dengan nama lembar kerja sumber Anda.
Set SH_PIVOT = twb.Sheets("tes3") ' Membuat lembar kerja baru untuk Pivot Table.

twb.Sheets("tes3").Select
' Loop melalui semua PivotTables di lembar kerja SH_PIVOT
For Each PivotTable In SH_PIVOT.PivotTables
        PivotTable.TableRange2.Clear
Next PivotTable

Set PivotRange = DataSrc.UsedRange ' Ganti dengan rentang yang sesuai.
Set PivotCache = twb.PivotCaches.Create(SourceType:=xlDatabase, _
                    SourceData:=PivotRange)
Set PivotTable = PivotCache.CreatePivotTable _
                    (TableDestination:=SH_PIVOT.Range("A1"), _
                    TableName:="PV_PerFactory")
                    
''' SETTING PIVOT '''
With PivotTable
    .DisplayErrorString = True
    .ErrorString = "-"
End With

'''' INSERT FIELD FIELD NYA ''''
With PivotTable.PivotFields("Bln.Prd")
    .Caption = "Bln.Prd"
    .Orientation = xlRowField
    .Position = 1
    .PivotItems("(blank)").Visible = False
End With
With PivotTable.PivotFields("Factory")
    .Caption = "Factory"
    .Orientation = xlRowField
    .Position = 2
    .PivotItems("(blank)").Visible = False
End With

With PivotTable.PivotFields("QTY")
    .Orientation = xlDataField
    .Position = 1
    .Function = xlSum
    .Caption = " QTY"
    .NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
End With
With PivotTable.PivotFields("Am.FOB")
    .Orientation = xlDataField
    .Position = 2
    .Function = xlSum
    .Caption = "Sum of AM.FOB"
    .NumberFormat = "_($* #,##0.00_);[Red]_($* (#,##0.00);_(* ""-""??_);_(@_)"
End With
With PivotTable.PivotFields("Amt.CM(USD)")
    .Orientation = xlDataField
    .Position = 3
    .Function = xlSum
    .Caption = "Sum of Am.CM"
    .NumberFormat = "_($* #,##0.00_);[Red]_($* (#,##0.00);_(* ""-""??_);_(@_)"
End With
With PivotTable.PivotFields("Cost.Line(USD)")
    .Orientation = xlDataField
    .Position = 4
    .Function = xlSum
    .Caption = "Sum of Cost Line"
    .NumberFormat = "_($* #,##0.00_);[Red]_($* (#,##0.00);_(* ""-""??_);_(@_)"
End With
With PivotTable.PivotFields("Profit.Lost(USD)")
    .Orientation = xlDataField
    .Position = 5
    .Function = xlSum
    .Caption = "Sum of Profit/Loss"
    .NumberFormat = "_($* #,##0.00_);[Red]_($* (#,##0.00);_(* ""-""??_);_(@_)"
End With
With PivotTable
    .CalculatedFields.Add "%", "='Profit.Lost(USD)' /Am.FOB", True
    .PivotFields("%").Orientation = xlDataField
    .PivotFields("Sum of %").Position = 6
    .PivotFields("Sum of %").Function = xlSum
    .PivotFields("Sum of %").Caption = "Sum of %"
    .PivotFields("Sum of %").NumberFormat = "0%;[Red]-0%;_(* """"-""""??_);_(@_)"
End With

' SETTING TABULAR & TAMPILAN
With PivotTable
    .RowAxisLayout xlTabularRow
    .PivotFields("Factory").Subtotals = Array _
        (False, False, False, False, False, False, False, False, False, False, False, False)
    .DataPivotField.Caption = " "
    .TableStyle2 = "PivotStyleMedium9"
End With

ActiveSheet.PivotTables("PV_PerFactory").PivotSelect _
    "'Sum of %'", xlDataAndLabel, True
With Selection
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
End With

' GANTI NAMA FACTORY

ActiveSheet.PivotTables("PV_PerFactory").PivotSelect _
    "Factory[All]", xlLabelOnly, True
With Selection
    Dim SearchValues As Variant
    Dim ReplaceValues As Variant
    Dim i As Long
    SearchValues = Array("CJL", "CHW", "KLB", "MJ1", "MJ2", "CVA")
    ReplaceValues = Array("CLN", "CHAWAN", "KALIBENDA", "MAJA 1", "MAJA 2", "KRAPYAK")
    
    For i = LBound(SearchValues) To UBound(SearchValues)
        Selection.Replace What:=SearchValues(i), Replacement:=ReplaceValues(i), LookAt:=xlPart, _
            SearchOrder:=xlByColumns, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Next i
End With

End Sub