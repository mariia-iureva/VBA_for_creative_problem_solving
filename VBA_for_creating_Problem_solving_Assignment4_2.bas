Option Explicit

Function prime(n As Integer) As Boolean
'Place your code here
Dim l As Integer, flag As Boolean, i As Integer
'Initializing
l = WorksheetFunction.RoundDown(Sqr(n), 0)
flag = True
For i = 2 To l
    If n Mod i = 0 Then flag = False
Next i
prime = flag
End Function

'Place your code here
Function countprime(n1 As Integer, n2 As Integer) As Integer
Dim i As Integer, c As Integer
For i = n1 To n2
    If prime(i) = True Then c = c + 1
Next i
countprime = c
End Function

