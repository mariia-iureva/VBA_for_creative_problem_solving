Option Explicit

Function tank(R As Double, H As Double, d As Double) As Double
Dim pi As Double
pi = WorksheetFunction.pi
'Place your code here
If d <= R Then
    tank = ((pi * (d ^ 2)) / 3) * (3 * R - d)
ElseIf R < d <= (H - R) Then
    tank = (2 * pi * R ^ 3) / 3 + pi * R ^ 2 * (d - R)
ElseIf (H - R) < d <= H Then
    tank = (4 * pi * R ^ 3) / 3 + pi * R ^ 2 * (H - 2 * R) - ((pi * (H - d) ^ 2) / 3) * (3 * R - H + d)
End If
End Function
