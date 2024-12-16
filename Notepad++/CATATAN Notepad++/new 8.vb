Sub GetKonker_ToIC()

    SH1_IC.Activate
    SH1_IC.AutoFilterMode = False
    SH1_IC.Range("F:F").ClearContents
    SH1_IC.Range("F7") = "KK"
    ' Tentukan baris terakhir di Sheet1 dan Sheet2
    
    LR1 = SH1_IC.Cells(SH1_IC.Rows.Count, "N").End(xlUp).Row
    LR1_2 = SH1_TO_PVT.Cells(SH1_TO_PVT.Rows.Count, "N").End(xlUp).Row

    ' Loop melalui setiap nilai di kolom N pada Sheet1
    For Each cell In SH1_IC.Range("N8:N" & LR1)
        If cell.Offset(0, -4) = "IC" Then
            lookupValue = cell.Value
            result = "" ' Reset result untuk setiap lookup value
    
            ' Loop melalui setiap cell di Sheet2 untuk mencari nilai yang cocok
            For Each rng In SH1_TO_PVT.Range("O2:O" & LR1_2)
                ' Jika nilai cocok, tambahkan ke hasil dengan pemisah "|"
                If rng.Value <> "" Then
                    If rng.Value = lookupValue Then
                        If result = "" Then
                            result = rng.Offset(0, -1).Value ' Ambil nilai di kolom sebelah kanan dari hasil yang cocok
                        Else
                            If rng.Offset(0, -1).Value <> Left(Results, 8) Then
                                result = result & " | " & rng.Offset(0, -1).Value
                            End If
                        End If
                    End If
                End If
            Next rng
    
            ' Simpan hasil di kolom F pada baris yang sama di Sheet1
            cell.Offset(0, -8).Value = result ' Sesuaikan Offset sesuai dengan posisi kolom F
        End If
    Next cell
End Sub