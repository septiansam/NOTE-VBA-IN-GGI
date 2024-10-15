Attribute VB_Name = "Module1"
'***Because Giving is Something, But Sharing is Everything!
'***Long Live OpenSource!!!

Public Function MD5Hash( _
  ByVal strText As String) _
  As String

' Create and return MD5 signature from strText.
' Signature has a length of 32 characters.
'
' 2005-11-21. Cactus Data ApS, CPH.

  Dim cMD5          As New clsMD5
  Dim strSignature  As String
  
  ' Calculate MD5 hash.
  strSignature = cMD5.MD5(strText)
  
  ' Return MD5 signature.
  MD5Hash = strSignature
  
  Set cMD5 = Nothing
  
End Function

Public Function IsMD5( _
  ByVal strText As String, _
  ByVal strMD5 As String) _
  As Boolean
  
' Checks if strMD5 is the MD5 signature of strText.
' Returns True if they match.
' Note: strText is case sensitive while strMD5 is not.
'
' 2005-11-21. Cactus Data ApS, CPH.

  Dim booMatch  As Boolean
  
  booMatch = (StrComp(strMD5, MD5Hash(strText), vbTextCompare) = 0)
  IsMD5 = booMatch
  
End Function

Sub btn_Convert()
    Dim name As String
    Dim Sep As String
    Dim Response As Integer

    Filename = Application.GetSaveAsFilename(InitialFileName:=vbNullString, fileFilter:="CSV Files (*.csv),*.csv")
    If Filename = False Then
        ''''''''''''''''''''''''''
        ' user cancelled, get out
        ''''''''''''''''''''''''''
        Exit Sub
    End If
    name = CStr(Filename)
    If FileExists(name) Then
        Beep
        Response = MsgBox(prompt:="File Already Exist, Do you want to replace it?", Buttons:=vbYesNo)
        If Response = vbYes Then
            ExportToTextFile FName:=name, Sep:=",", _
               SelectionOnly:=False, AppendData:=False
        Else: Exit Sub
        End If
    Else
        If Right(name, 4) = ".txt" Then
                name = Mid(name, 1, Len(name) - 4)
                name = name & ".csv"
        End If
        ExportToTextFile FName:=name, Sep:=",", _
               SelectionOnly:=False, AppendData:=False
    End If
End Sub

Function FileExists(stFile As String) As Boolean
    'Uses FULL filename
    If Dir(stFile) <> "" Then FileExists = True
End Function
