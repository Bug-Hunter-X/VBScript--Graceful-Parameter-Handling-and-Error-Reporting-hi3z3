Function MyFunction(param1, param2)
  If IsEmpty(param1) Or IsEmpty(param2) Then
    Err.Raise 13, , "Both parameters are required."
  End If
  ' ... rest of the function ...
End Function