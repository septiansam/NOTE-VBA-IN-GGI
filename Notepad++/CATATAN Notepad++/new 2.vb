Sub CreatePivotReport()
    SH1_TO_PVT.Activate
    LR1 = SH1_TO_PVT.Range("A" & Rows.Count).End(xlUp).Row
    
    Set pv_Rng = Range("A1:K" & LR1)
    Set pv_Cache = WB1.PivotCaches.Create(SourceType:=xlDatabase, _
                        SourceData:=pv_Rng)
    Set pv_Tb = pv_Cache.CreatePivotTable _
                        (TableDestination:=SH1_PVT.Range("A1"))
    SH1_PVT.Activate
    
    With pv_Tb
        .RowAxisLayout xlTabularRow
        .RowGrand = False
        .NullString = ""
        .ColumnGrand = False
'        .ShowValuesRow = False
        .DisplayErrorString = False
        .RepeatAllLabels xlRepeatLabels
        .PivotCache.MissingItemsLimit = xlMissingItemsDefault
    End With
    
    With pv_Tb.PivotFields("BU")
        .Caption = "BU"
        .Orientation = xlRowField
        .Position = 1
        .Subtotals = _
            Array(False, False, False, False, False, False, False, False, False, False, False, False)
    End With
    
    With pv_Tb.PivotFields("KONKER")
        .Caption = "KONKER"
        .Orientation = xlRowField
        .Position = 2
'        .Subtotals = _
'            Array(False, False, False, False, False, False, False, False, False, False, False, False)
    End With
    
    With pv_Tb.PivotFields("WO")
        .Caption = "WO"
        .Orientation = xlRowField
        .Position = 3
        .Subtotals = _
            Array(False, False, False, False, False, False, False, False, False, False, False, False)
    End With
    
    With pv_Tb.PivotFields("ITEM")
        .Caption = "ITEM"
        .Orientation = xlRowField
        .Position = 4
        .Subtotals = _
            Array(False, False, False, False, False, False, False, False, False, False, False, False)
    End With
    
    With pv_Tb.PivotFields("UOM")
        .Caption = "UOM"
        .Orientation = xlRowField
        .Position = 5
        .Subtotals = _
            Array(False, False, False, False, False, False, False, False, False, False, False, False)
    End With
    
    With pv_Tb.PivotFields("SOURCE1")
        .Orientation = xlColumnField
        .Position = 1
    End With
    
    With pv_Tb.PivotFields("QTY")
        .Orientation = xlDataField
        .Position = 1
        .Function = xlSum
        .NumberFormat = "_(* #,##0.00_);[Red]_(* (#,##0.00);_(* ""-""??_)"
    End With
    
End Sub

Sub CreatePivot_ForGetKonker()
    SH1_TO_PVT.Activate
    
    SH1_TO_PVT.Range("N:O").ClearContents
    
    LR1 = SH1_TO_PVT.Range("A" & Rows.Count).End(xlUp).Row
    Set pv_Rng = Range("I1:J" & LR1)
    Set pv_Cache = WB1.PivotCaches.Create(SourceType:=xlDatabase, _
                        SourceData:=pv_Rng)
    Set pv_Tb = pv_Cache.CreatePivotTable _
                        (TableDestination:=SH1_TO_PVT.Range("N1"))
    
    With pv_Tb
        .RowAxisLayout xlTabularRow
        .RowGrand = False
        .NullString = ""
        .ColumnGrand = False
'        .ShowValuesRow = False
        .DisplayErrorString = False
        .RepeatAllLabels xlRepeatLabels
        .PivotCache.MissingItemsLimit = xlMissingItemsDefault
    End With
    
    With pv_Tb.PivotFields("KONKER")
        .Caption = "KONKER"
        .Orientation = xlRowField
        .Position = 1
        .Subtotals = _
            Array(False, False, False, False, False, False, False, False, False, False, False, False)
    End With
    
    With pv_Tb.PivotFields("WO")
        .Caption = "WO"
        .Orientation = xlRowField
        .Position = 2
        .PivotItems("(blank)").Visible = False
        .Subtotals = _
            Array(False, False, False, False, False, False, False, False, False, False, False, False)
    End With
    
    LR1 = SH1_TO_PVT.Range("N" & Rows.Count).End(xlUp).Row
    With SH1_TO_PVT.Range("N1:O" & LR1)
        .Copy: .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
    End With
    
    SH1_TO_PVT.Cells(1, 1).Select

End Sub