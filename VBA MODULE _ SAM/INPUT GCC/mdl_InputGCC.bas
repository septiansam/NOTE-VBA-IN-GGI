Attribute VB_Name = "mdl_InputGCC"



    Dim WB1 As ThisWorkbook
    Dim SH1_DB As Worksheet, SH1_Bantuan As Worksheet
    Dim LR1 As Long
    Dim str_PathInputGCC As String
    '-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_
    Dim WB2_InputGCC As Workbook
    Dim SH2 As Worksheet
    

Sub BUTTON_InputData_ToGCC()
    Application.AskToUpdateLinks = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Call Validasi_Data
    Call Main_Processing

    Application.AskToUpdateLinks = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
End Sub

Sub Main_Processing()
    Set WB1 = ThisWorkbook
    Set SH1_DB = WB1.Worksheets("Database Supplier")
    Set SH1_Bantuan = WB1.Worksheets("Bantuan")
    
    SH1_DB.Activate
    str_PathInputGCC = SH1_DB.Range("L13").Value
    SH1_Bantuan.Activate
    SH1_Bantuan.Copy
    Set WB2_InputGCC = ActiveWorkbook
    Windows(WB2_InputGCC.Name).Activate
    Set SH2 = WB2_InputGCC.Worksheets(1)
    SH2.Activate
    SH2.Name = "Input"
    SH2.Range("B:B").Delete
    SH2.Range("C:G").Delete
    SH2.Range("D:D").Delete
    SH2.Range("E1").Value = "KURS"
    SH2.Range("F1").Value = "Status"
    
    SH2.Rows(1).Font.Bold = True
    SH2.Cells.EntireColumn.AutoFit
    SH2.Cells(1, 1).Select
    
    WB2_InputGCC.SaveAs str_PathInputGCC, xlOpenXMLWorkbook
    WB2_InputGCC.Close True
    Set WB2_InputGCC = Nothing
    Set SH2 = Nothing
    
    Windows(WB1.Name).Activate
    SH1_DB.Activate
    SH1_DB.Cells(1, 1).Select
End Sub

Sub Validasi_Data()

    Set WB1 = ThisWorkbook
    Set SH1_DB = WB1.Worksheets("Database Supplier")

    SH1_DB.Activate
    str_PathInputGCC = SH1_DB.Range("L13").Value
    SH1_DB.Range("L13").Hyperlinks.Delete
    If Dir(str_PathInputGCC) = "" Then
        MsgBox "File Input GCC Tidak Ditemukan", vbCritical, "FILE NOT FOUND"
        SH1_DB.Activate
        SH1_DB.Cells(1, 1).Select
        End
    End If
    SH1_DB.Activate
    SH1_DB.Cells(1, 1).Select
End Sub
