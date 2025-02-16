Function f(a,b)
  If a>b then
    MsgBox "a is greater than b"
  ElseIf a<b then
    MsgBox "a is less than b"
  Else
    MsgBox "a is equal to b"
  End If
end function

'This function will not work correctly if a or b are null or empty strings.
'The if statements will result in a type mismatch error.

f(1,"") 'This will cause a type mismatch error because an integer is being compared to a string.
f(null,2) 'This will cause a type mismatch error because null is not a number.
