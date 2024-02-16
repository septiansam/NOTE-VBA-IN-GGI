Attribute VB_Name = "Module3"
'nameshee
'output

Sub ExampleLoop()
    Dim i As Integer
    Dim total As Integer
    
    ' Looping dengan For Next
    For i = 1 To 10
        Debug.Print "For Next Loop: " & i
    Next i
    
    ' Looping dengan Do While
    i = 1
    Do While i <= 5
        Debug.Print "Do While Loop: " & i
        i = i + 1
    Loop
    
    ' Looping dengan Do Until
    i = 1
    Do Until i > 5
        Debug.Print "Do Until Loop: " & i
        i = i + 1
    Loop
    
    ' Looping dengan Do Loop While
    i = 1
    Do
        Debug.Print "Do Loop While Loop: " & i
        i = i + 1
    Loop While i <= 5
    
    ' Looping dengan Do Loop Until
    i = 1
    Do
        Debug.Print "Do Loop Until Loop: " & i
        i = i + 1
    Loop Until i > 5
    
    ' Looping dengan For Each Next
    Dim names As Variant
    names = Array("John", "Jane", "Bob", "Alice")
    For Each Name In names
        Debug.Print "For Each Next Loop: " & Name
    Next Name
    
    ' Looping dengan Exit For
    total = 0
    For i = 1 To 10
        total = total + i
        If total >= 15 Then
            Exit For
        End If
    Next i
    
    Debug.Print "Total: " & total
End Sub

