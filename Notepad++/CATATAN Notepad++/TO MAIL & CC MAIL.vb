'''[##].. GET TO EMAIL
    SH1_OLAH1.Cells.ClearContents
    SH1_RefMail.Activate
    LR1 = SH1_RefMail.Range("A" & Rows.Count).End(xlUp).Row
    SH1_RefMail.Range("A2:A" & LR1).Copy
    SH1_OLAH1.Activate
    SH1_OLAH1.Range("A1").PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
    SH1_OLAH1.Range("A:A").RemoveDuplicates 1, xlNo
    LR1 = SH1_OLAH1.Range("A" & Rows.Count).End(xlUp).Row
    Str_ToEmail = ""
    For i = 1 To LR1
        Str_ToEmail = Str_ToEmail & SH1_OLAH1.Range("A" & i).Value & ","
    Next i
    Str_ToEmail = VBA.Left(Str_ToEmail, Len(Str_ToEmail) - 1)
    
    '''[##].. GET CC EMAIL
    SH1_OLAH1.Cells.ClearContents
    SH1_RefMail.Activate
    LR1 = SH1_RefMail.Range("D" & Rows.Count).End(xlUp).Row
    If LR1 > 1 Then
        SH1_RefMail.Range("D2:D" & LR1).Copy
        SH1_OLAH1.Activate
        SH1_OLAH1.Range("A1").PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
        SH1_OLAH1.Range("A:A").RemoveDuplicates 1, xlNo
        LR1 = SH1_OLAH1.Range("A" & Rows.Count).End(xlUp).Row
        Str_CCEmail = ""
        For i = 1 To LR1
            Str_CCEmail = Str_CCEmail & SH1_OLAH1.Range("A" & i).Value & ","
        Next i
        Str_CCEmail = VBA.Left(Str_CCEmail, Len(Str_CCEmail) - 1)
    Else
        Str_CCEmail = ""
    End If
    
    '[*].. KE SHEETS RPA4
    SH1_RPA4.Activate
    SH1_RPA4.Range("A1").Value = "TO"
    SH1_RPA4.Range("A2").Value = Str_ToEmail
    
    SH1_RPA4.Range("B1").Value = "CC"
    SH1_RPA4.Range("B2").Value = Str_CCEmail