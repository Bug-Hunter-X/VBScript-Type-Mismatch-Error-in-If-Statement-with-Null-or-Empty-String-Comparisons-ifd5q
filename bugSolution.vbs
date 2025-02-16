Function f(a,b)
  If IsNumeric(a) And IsNumeric(b) Then
    If a > b Then
      MsgBox "a is greater than b"
    ElseIf a < b Then
      MsgBox "a is less than b"
    Else
      MsgBox "a is equal to b"
    End If
  ElseIf IsEmpty(a) Or IsEmpty(b) Then
    MsgBox "One or both input values are empty"
  ElseIf a = "" Or b = "" Then
    MsgBox "One or both input values are empty strings"
  Else
    MsgBox "Invalid input type: Both inputs must be numeric"
  End If
end function

'The improved function now explicitly checks for numeric values
'using IsNumeric() and handles empty strings and null values appropriately.

f(1,"")
f(null,2)
f(1,2)
f(3,1) 