Function MyFunction(param1, param2)
  If IsEmpty(param1) Then
    Err.Raise vbError + 1, , "Parameter 1 is required."
  ElseIf IsEmpty(param2) Then
    Err.Raise vbError + 2, , "Parameter 2 is required."
  Else
    ' ... rest of the function ...
  End If
End Function

'Improved error handling to provide more specific information for each missing parameter.