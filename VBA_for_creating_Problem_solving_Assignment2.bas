Attribute VB_Name = "Module1"
Option Explicit

Function antoine(A As Double, B As Double, C As Double, t As Double) As Double
antoine = 10 ^ (A - (B / (t + C)))
'Place your code here
End Function

Function medication(c0 As Double, k As Double, t As Double) As Double
'Place your code here
medication = c0 * Exp(-k * t)
End Function

Function payment(P As Double, i As Double, n As Double) As Double
'Place your code here
payment = (P * i / 12) / (1 - (1 + i / 12) ^ ((-n) * 12))
End Function
