Function MyFunction(param)
  'Early Binding Example
  If TypeName(param) = "Double" Then
    result = param * 2
  Else
    Err.Raise vbError, , "Parameter must be a number."
  End If
  MyFunction = result
End Function

Sub Main()
  On Error Resume Next
  Dim x, result
  x = 5 'Correct Type
  result = MyFunction(x)
  If Err.Number <> 0 Then
     MsgBox "Error: " & Err.Description
     Err.Clear
  End If
  MsgBox "Result: " & result
  
  x = "hello" 'Incorrect Type
  result = MyFunction(x)
  If Err.Number <> 0 Then
     MsgBox "Error: " & Err.Description
     Err.Clear
  End If
  MsgBox "Result: " & result
  
  On Error GoTo 0
End Sub