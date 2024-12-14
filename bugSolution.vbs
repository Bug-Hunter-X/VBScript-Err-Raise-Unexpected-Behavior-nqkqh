Function MyFunc(param)
  On Error GoTo ErrorHandler
  If IsEmpty(param) Then
    Err.Raise 13, , "Parameter cannot be empty"
  End If
  ' ... rest of the function
  Exit Function
ErrorHandler:
  ' Handle the error gracefully
  MsgBox "Error: " & Err.Description
  Err.Clear
End Function