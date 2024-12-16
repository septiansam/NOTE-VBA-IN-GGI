Sub GetFileNamesAndExtensions()
    Dim fso As Object
    Dim objFolder As Object
    Dim objFile As Object
    Dim i As Integer

    ' Ganti "C:\Folder\Path" dengan path folder yang ingin Anda periksa
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set objFolder = fso.GetFolder("C:\Folder\Path")

    i = 1

    ' Loop melalui setiap file dalam folder
    For Each objFile In objFolder.files
        ' Mengambil nama file dan ekstensinya
        Cells(i, 1).Value = objFile.name
        Cells(i, 2).Value = fso.GetExtensionName(objFile.Path)
        i = i + 1
    Next objFile

    ' Membersihkan objek yang digunakan
    Set objFolder = Nothing
    Set fso = Nothing
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
	Application.DisplayAlerts = False
	Application.ScreenUpdating = False
	Application.EnableEvents = False

	Dim affectedRange As Range, cell As Range, RNG_DATE As Range
	Set affectedRange = Intersect(Target, Me.Range("J2:J100000"))

	If Not affectedRange Is Nothing Then
		For Each cell In affectedRange
			If cell.value <> "" Then
				cell.Offset(0, 1).Formula = "=IFERROR(INDEX(MASTER!D:D,MATCH(" & cell.Address & ",MASTER!E:E,0)),"""")"
			ElseIf cell.value = vbNullString Then
				cell.Offset(0, 1).value = vbNullString
			End If
		Next cell
	End If

	Dim STR_DATE As String
	Dim DATE_VALUE As Date

	Set RNG_DATE = Intersect(Target, Me.Range("B2:B100000"))
	If Not RNG_DATE Is Nothing And Target.Cells.Count = 1 Then
		For Each cell In RNG_DATE
			If cell.value <> "" And _
				IsDate(cell.value) And _
				(Format(cell.value, "DD/MM/YYYY") = cell.Text Or Format(cell.value, "D/M/YYYY") = cell.Text) Then
				
				STR_DATE = Format(Target.value, "M/D/YYYY")
				Target.value = STR_DATE
			End If
		Next cell
	End If

	Sheets("RPA").Select
	'Cells(1, 1).Select

	Application.DisplayAlerts = True
	Application.ScreenUpdating = True
	Application.EnableEvents = True
End Sub


Private Sub Worksheet_Change(ByVal Target As Range)
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.EnableEvents = False

Dim affectedRange As Range, Cell As Range
Set affectedRange = Intersect(Target, Union(Me.Range("E2:E100000"), Me.Range("K2:K100000")))

If Not affectedRange Is Nothing Then
    For Each Cell In affectedRange
        If Cell.Column = 5 Then
            If Cell.Value <> vbNullString Then
                Cell.Offset(0, 1).Formula = "=IFERROR(INDEX(ADDRESS!A:A,MATCH(INPUT!" & Cell.Address & ",ADDRESS!B:B,0)),""Data Placing Tidak Ditemukan, Mohon Klik Tombol Di Sheets TOMBOL"")"
                If Cell.Offset(0, 1).Value = "Data Placing Tidak Ditemukan, Mohon Klik Tombol Di Sheets TOMBOL" Then
                    With Cell.Offset(0, 1)
                        .Font.Bold = True
                        .Font.Color = vbRed
                    End With
                Cell.EntireColumn.AutoFit
                End If
            ElseIf Cell.Value = vbNullString Then
                Cell.Offset(0, 1).Value = vbNullString
                With Cell.Offset(0, 1)
                    .Font.Bold = False
                    .Font.Color = vbBlack
                    .EntireColumn.AutoFit
                End With
            End If
        ElseIf Cell.Column = 11 Then
            If Cell.Value <> "" Then
                Cell.Offset(0, 1).Formula = "=IFERROR(VLOOKUP(" & Cell.Address & ",INMK!A:B,2,0),""Data ITEM MAKLOON Tidak Ditemukan, Mohon Klik Tombol Di Sheets TOMBOL"")"
                Cell.Offset(0, 2).Formula = "=IFERROR(VLOOKUP(" & Cell.Address & ",INMK!A:E,5,0),""Data ITEM MAKLOON Tidak Ditemukan, Mohon Klik Tombol Di Sheets TOMBOL"")"
                Cell.Offset(0, 3).Formula = "=IFERROR(VLOOKUP(" & Cell.Address & ",INMK!A:G,7,0),""Data ITEM MAKLOON Tidak Ditemukan, Mohon Klik Tombol Di Sheets TOMBOL"")"
                If Cell.Offset(0, 1).Value = "Data ITEM MAKLOON Tidak Ditemukan, Mohon Klik Tombol Di Sheets TOMBOL" Or _
                    Cell.Offset(0, 2).Value = "Data ITEM MAKLOON Tidak Ditemukan, Mohon Klik Tombol Di Sheets TOMBOL" Then
                    With Cell.Offset(0, 1)
                        .Font.Bold = True
                        .Font.Color = vbRed
                        .EntireColumn.AutoFit
                    End With
                    With Cell.Offset(0, 2)
                        .Font.Bold = True
                        .Font.Color = vbRed
                        .EntireColumn.AutoFit
                    End With
                Else
                    With Cell.Offset(0, 1)
                        .Font.Bold = False
                        .Font.Color = vbBlack
                        .EntireColumn.AutoFit
                    End With
                    With Cell.Offset(0, 2)
                        .Font.Bold = False
                        .Font.Color = vbBlack
                        .EntireColumn.AutoFit
                    End With
                End If
            ElseIf Cell.Value = vbNullString Then
                Cell.Offset(0, 1).Value = vbNullString
                Cell.Offset(0, 2).Value = vbNullString
                With Cell.Offset(0, 1)
                    .Font.Bold = False
                    .Font.Color = vbBlack
                    .EntireColumn.AutoFit
                End With
                With Cell.Offset(0, 2)
                    .Font.Bold = False
                    .Font.Color = vbBlack
                    .EntireColumn.AutoFit
                End With
            End If
        End If
    Next Cell
End If

Application.DisplayAlerts = True
Application.ScreenUpdating = True
Application.EnableEvents = True
End Sub