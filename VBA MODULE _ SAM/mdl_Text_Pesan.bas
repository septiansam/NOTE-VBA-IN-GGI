Attribute VB_Name = "mdl_Text_Pesan"


Sub Text_Email()
    Dim formattedText As String
    Dim str_Waktu As String, str_Pesan1 As String
    Dim currentTime As Double
    
    'Mendapatkan waktu saat ini
    currentTime = Time
    Select Case currentTime
        Case TimeValue("05:00:00") To TimeValue("10:59:59")
            str_Waktu = "Pagi"
        Case TimeValue("11:00:00") To TimeValue("14:59:59")
            str_Waktu = "Siang"
        Case TimeValue("15:00:00") To TimeValue("18:00:00")
            str_Waktu = "Sore"
        Case Else
            str_Waktu = "Malam"
    End Select
    
    str_Pesan1 = ThisWorkbook.Worksheets("RPA4").Range("C2").Value
    
    formattedText = "Selamat " & str_Waktu & "," & vbLf & _
                    "Berikut adalah " & str_Pesan1 & vbLf & _
                    "(File Terlampir)" & vbLf & _
                    "Terimakasih" & vbCrLf & vbLf & _
                    "GISCA" & vbLf & _
                    "Gistex Communication Assistant" & vbLf & _
                    "Please do not repay on this number"
    
    ThisWorkbook.Worksheets("RPA4").Range("E2").Value = formattedText
    ThisWorkbook.Worksheets("RPA4").Range("E2").WrapText = False
End Sub

Sub Text_WA()
    Dim formattedText As String
    Dim str_Waktu As String, str_Pesan1 As String
    Dim currentTime As Double
    
    'Mendapatkan waktu saat ini
    currentTime = Time
    Select Case currentTime
        Case TimeValue("05:00:00") To TimeValue("10:59:59")
            str_Waktu = "Pagi"
        Case TimeValue("11:00:00") To TimeValue("14:59:59")
            str_Waktu = "Siang"
        Case TimeValue("15:00:00") To TimeValue("18:00:00")
            str_Waktu = "Sore"
        Case Else
            str_Waktu = "Malam"
    End Select
    
    str_Pesan1 = ThisWorkbook.Worksheets("RPA4").Range("C2").Value
    
    formattedText = "Selamat " & str_Waktu & "," & vbLf & _
                    "Berikut adalah " & str_Pesan1 & vbLf & _
                    "(File Terlampir)" & vbLf & _
                    "Terimakasih" & vbCrLf & vbLf & _
                    "GISCA" & vbLf & _
                    "*Gistex Communication Assistant*" & vbLf & _
                    "_Please do not repay on this number_"
    
    ThisWorkbook.Worksheets("RPA4").Range("E2").Value = formattedText
    ThisWorkbook.Worksheets("RPA4").Range("E2").WrapText = False
End Sub
