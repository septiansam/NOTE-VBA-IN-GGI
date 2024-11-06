Attribute VB_Name = "mdl_Text_GetDriveSpace"
Sub GetDriveSpace()
    Dim fso As Object
    Dim drive As Object
    Dim totalSpace As Double
    Dim freeSpace As Double
    Dim usedSpace As Double

    ' Membuat instance FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Mendapatkan drive Z:
    Set drive = fso.GetDrive("Z:\")
    
    ' Mengambil Total Space dan Free Space dalam GB
    totalSpace = drive.TotalSize / 1024 / 1024 / 1024  ' Konversi ke GB
    freeSpace = drive.freeSpace / 1024 / 1024 / 1024   ' Konversi ke GB
    usedSpace = totalSpace - freeSpace                 ' Menghitung Used Space
    
    ' Menampilkan hasil
    MsgBox "Drive Z:\" & vbCrLf & _
           "Total Space: " & Format(totalSpace, "0.00") & " GB" & vbCrLf & _
           "Free Space: " & Format(freeSpace, "0.00") & " GB" & vbCrLf & _
           "Used Space: " & Format(usedSpace, "0.00") & " GB"

    ' Membersihkan objek
    Set drive = Nothing
    Set fso = Nothing
End Sub

