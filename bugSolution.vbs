Function MyFunction(param1, param2)
  On Error GoTo ErrHandler
  If IsEmpty(param1) Or IsEmpty(param2) Then
    Err.Raise vbErrorArgumentNotSupplied, , "Parameters cannot be empty"
  End If
  ' ...rest of the function...
  Exit Function
ErrHandler:
  MsgBox "An error occurred: " & Err.Description, vbCritical
  Err.Clear
End Function