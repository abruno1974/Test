Function isEvenOdd(num)
if (num Mod 2) then
isEvenOdd = "Even"
else
isEvenOdd="Odd"
End if
End Function

Dim num
n=41
MsgBox n & "is" & isEvenOdd(n)