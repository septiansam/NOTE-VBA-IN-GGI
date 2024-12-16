Sub Find_OR()

Application.DisplayAlerts = False

Dim lastRow As Long, x As Long, i As Long, j As Long
Dim dataFind As String, dataPO As String
Dim found As Boolean
Dim candidate As String

' Set worksheet tempat Anda bekerja, ganti "Sheet1" dengan nama sheet Anda
Set ws = ThisWorkbook.Sheets(1)

lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

Range("A:A").Insert: Range("A1") = "BU"
With Range("A2:A" & lastRow)
    .FormulaR1C1 = "=C3*1"
    .Copy
    .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False: Cells(1, 1).Select
End With

Range("A:A").Insert: Range("A1") = "OR"

' Loop melalui semua baris

For i = 2 To lastRow
    dataFind = ws.Cells(i, 9).Value ' Kolom A
    dataFind = Trim$(dataFind)
    dataPO = ws.Cells(i, 10).Value ' Kolom B
    found = False

    For j = Len(dataFind) - 7 To 1 Step -1
        If Left(dataFind, 2) Like String(2, "#") Or _
            Left(dataFind, 1) Like String(1, "/") Then

            candidate = Mid(dataFind, j, 8)
            If candidate Like String(8, "#") _
                And candidate <> dataPO _
                And Right(candidate, 6) <> Right(dataPO, 6) _
                And Left(candidate, 1) = Left(dataPO, 1) Then
                
                ws.Cells(i, 1).Value = candidate
                found = True
                Exit For ' Keluar dari loop jika sudah menemukan
            End If
        End If
    Next j
    
    If Not found Then
        ws.Cells(i, 1).Value = "Tidak ada"
    End If
Next i

Application.DisplayAlerts = True

End Sub
