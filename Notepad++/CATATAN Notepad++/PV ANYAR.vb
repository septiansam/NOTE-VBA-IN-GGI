Option Explicit

Sub CreatePivot()

    reportPO.Activate
    
    lr = Cells.Find("*", , xlFormulas, xlPart, xlByRows, xlPrevious).Row
    lc = Cells.Find("*", , xlFormulas, xlPart, xlByColumns, xlPrevious).Column

    Set rgPv = Range(Cells(5, 1), Cells(lr, lc))
    Set pvCh = TWB.PivotCaches.Create(SourceType:=xlDatabase, _
                        SourceData:=rgPv)
    Set pvTb = pvCh.CreatePivotTable _
                        (TableDestination:=temp.Range("A1"), _
                        TableName:="pv_table")
    
    temp.Activate
    With pvTb.PivotFields("EX FTY")
        .Caption = "EX FTY"
        .Orientation = xlRowField
        .Position = 1
        .NumberFormat = "m/d/yyyy"
    End With
    
    With pvTb.PivotFields("ORD NO")
        .Caption = "ORD NO"
        .Orientation = xlRowField
        .Position = 2
        .Subtotals = _
            Array(False, False, False, False, False, False, False, False, False, False, False, False)
    End With
    
    With pvTb.PivotFields("PO GARMENT")
        .Caption = "PO GARMENT"
        .Orientation = xlRowField
        .Position = 3
        .Subtotals = _
            Array(False, False, False, False, False, False, False, False, False, False, False, False)
    End With
    
    With pvTb.PivotFields("SO NO")
        .Caption = "SO NO"
        .Orientation = xlRowField
        .Position = 4
        .Subtotals = _
            Array(False, False, False, False, False, False, False, False, False, False, False, False)
    End With
    
    With pvTb.PivotFields("PLACING")
        .Caption = "PLACING"
        .Orientation = xlRowField
        .Position = 5
        .Subtotals = _
            Array(False, False, False, False, False, False, False, False, False, False, False, False)
    End With
    
    With pvTb.PivotFields("DESCRIPTION 1")
        .Caption = "DESCRIPTION 1"
        .Orientation = xlRowField
        .Position = 6
        .Subtotals = _
            Array(False, False, False, False, False, False, False, False, False, False, False, False)
    End With
    
    With pvTb.PivotFields("DESCRIPTION 2")
        .Caption = "DESCRIPTION 2"
        .Orientation = xlRowField
        .Position = 7
        .Subtotals = _
            Array(False, False, False, False, False, False, False, False, False, False, False, False)
    End With
    
    With pvTb.PivotFields("ORDER")
        .Orientation = xlDataField
        .Position = 1
        .Function = xlSum
    End With
    
    With pvTb.PivotFields("RECEIVED")
        .Orientation = xlDataField
        .Position = 2
        .Function = xlSum
    End With
    
    With pvTb.PivotFields("BALANCE")
        .Orientation = xlDataField
        .Position = 3
        .Function = xlSum
    End With
    
    With pvTb
        .RowAxisLayout xlTabularRow
        .ShowValuesRow = False
        .DisplayErrorString = False
    End With
    
End Sub

