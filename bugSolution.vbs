Function MyFunction(param1, param2)
  On Error GoTo ErrHandler
  If IsEmpty(param1) Or IsEmpty(param2) Then
    Err.Raise 1001, , "Parameters cannot be empty"
  End If
  ' ...rest of function code...
  Exit Function
ErrHandler:
  MsgBox "Error: " & Err.Number & " - " & Err.Description, vbCritical
End Function