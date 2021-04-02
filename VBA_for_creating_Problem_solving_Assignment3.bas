Attribute VB_Name = "Module1"
Option Explicit

Sub AddNumbersA()
'Place your code here
Dim x As Double
x = InputBox("Please, enter a number:")
Range("G12") = Range("D4") + x

End Sub

Sub AddNumbersB()
'Place your code here
Dim z As Double
z = InputBox("Please, enter a number:")
ActiveCell.Offset(-3, 2) = ActiveCell.Value + z
End Sub

Sub WherePutMe()
'Place your code here
Dim x As Double, letter As String
x = InputBox("Enter a row number:")
letter = InputBox("Enter a colon letter:")
Range(letter & x) = Selection.Cells(2, 2)
End Sub

Sub Swap()
'Place your code here
Dim x As Double, y As Double
x = Selection.Cells(1, 1)
y = Selection.Cells(1, 2)
Selection.Cells(1, 1) = y
Selection.Cells(1, 2) = x
End Sub
