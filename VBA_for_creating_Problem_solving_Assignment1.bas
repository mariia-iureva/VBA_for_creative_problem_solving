Attribute VB_Name = "Module1"
Option Explicit

Public RowsCount As Long
Public ColumnsCount As Long
Public row As Long

Public Col_DistrictTTL As Long
Public Col_SchoolTTL As Long
Public Col_HighestAchievement  As Long
Public Col_YearAfterGrad  As Long
Public Col_DemographicGrouping As Long
Public Col_DemographicValue As Long
Public Col_MedianEarnings As Long
Public Col_NumRecords As Long

Sub AdaResult()
    InitCol
        
    Dim result As String
    result = "Q1:" & vbNewLine & Q1 & vbNewLine
    result = result & vbNewLine & "Q2:" & vbNewLine & Q2 & vbNewLine
    result = result & vbNewLine & "Q3:" & vbNewLine & Q3 & vbNewLine
    result = result & vbNewLine & "Q4:" & Q4 & vbNewLine
    result = result & vbNewLine & "Q5:" & Q5 & vbNewLine
    
    MsgBox (result)
End Sub

Sub InitCol()
    Dim col As Long
    RowsCount = Cells(Rows.Count, 1).End(xlUp).row
    ColumnsCount = Cells(1, Columns.Count).End(xlToLeft).Column

    For col = 1 To ColumnsCount
        If Cells(1, col).Value = "DistrictTTL" Then
            Col_DistrictTTL = col
        ElseIf Cells(1, col).Value = "SchoolTTL" Then
            Col_SchoolTTL = col
        ElseIf Cells(1, col).Value = "HighestAchievement" Then
            Col_HighestAchievement = col
        ElseIf Cells(1, col).Value = "YearAfterGrad" Then
            Col_YearAfterGrad = col
        ElseIf Cells(1, col).Value = "DemographicGrouping" Then
            Col_DemographicGrouping = col
        ElseIf Cells(1, col).Value = "DemographicValue" Then
            Col_DemographicValue = col
        ElseIf Cells(1, col).Value = "MedianEarnings" Then
            Col_MedianEarnings = col
        ElseIf Cells(1, col).Value = "NumRecords" Then
            Col_NumRecords = col
        End If
    Next col
End Sub

Function Q1() As String
    Dim RowNumber As Long, Q1Range As Range, MaxRange As Range, MaxValue As Long, MaxRow As Long
    Set Q1Range = Range(Cells(2, Col_NumRecords), Cells(RowsCount, Col_NumRecords))
    MaxValue = WorksheetFunction.Max(Q1Range)
    Set MaxRange = Q1Range.Find(what:=MaxValue)
    MaxRow = MaxRange.row
    'RowNumber = Q1Range.Find(WorksheetFunction.Max(Q1Range).Row
    Q1 = Cells(MaxRow, Col_SchoolTTL).Value & " and " & Cells(MaxRow, Col_HighestAchievement).Value
End Function

Function Q2() As Long
    Q2 = 0
    For row = 2 To RowsCount
        If Cells(row, Col_HighestAchievement).Value = "Bachelor's or Higher" Then
            If Cells(row, Col_DemographicGrouping).Value = "All Students" Then
                Q2 = Q2 + Cells(row, Col_NumRecords).Value
            End If
        End If
    Next row
End Function

Function Q3() As Double
    Dim RangeForMedian() As Double
    Dim RangeCount As Long
    
    RangeCount = 0
    
    For row = 2 To RowsCount
        If Cells(row, Col_DemographicValue).Value = "FRPL" Then
            If Cells(row, Col_HighestAchievement).Value = "HS Diploma" Then
                ReDim Preserve RangeForMedian(RangeCount)
                RangeForMedian(RangeCount) = Cells(row, Col_MedianEarnings).Value
                RangeCount = RangeCount + 1
            End If
        End If
    Next row
    
    Q3 = Application.WorksheetFunction.Median(RangeForMedian)
End Function

Function Q4() As String

    Dim MaleRange() As Double, MaleCount As Long, MaleMaxValue As Long
    Dim FemaleRange() As Double, FemaleCount As Long, FemaleMaxValue As Long
    'InitCol
    
    For row = 2 To RowsCount
        If Cells(row, Col_HighestAchievement).Value = "Bachelor's or Higher" Then
            If Cells(row, Col_DemographicValue).Value = "Male" Then
                ReDim Preserve MaleRange(MaleCount)
                MaleRange(MaleCount) = Cells(row, Col_MedianEarnings).Value
                MaleCount = MaleCount + 1
            ElseIf Cells(row, Col_DemographicValue).Value = "Female" Then
                ReDim Preserve FemaleRange(FemaleCount)
                FemaleRange(FemaleCount) = Cells(row, Col_MedianEarnings).Value
                FemaleCount = FemaleCount + 1
            End If
        End If
    Next row
            
            MaleMaxValue = Application.WorksheetFunction.Max(MaleRange)
            FemaleMaxValue = Application.WorksheetFunction.Max(FemaleRange)
    
    Q4 = "Male Max Median is $" & MaleMaxValue - FemaleMaxValue & " or " & Round(MaleMaxValue / FemaleMaxValue, 2) & " times or " & Round((MaleMaxValue - FemaleMaxValue) * 100 / FemaleMaxValue, 0) & "% higher than Female Max Median"
    
End Function

Function Q5() As String
    'InitCol
    Dim Districts() As String, DistrictCount As Long, DistrictName As String
    Dim MedianHouseholdIncome2019 As Long
    
    MedianHouseholdIncome2019 = 73775
     
    For row = 2 To RowsCount
        If Cells(row, Col_MedianEarnings).Value > MedianHouseholdIncome2019 Then
            If Cells(row, Col_DistrictTTL).Value <> DistrictName Then
                DistrictName = Cells(row, Col_DistrictTTL).Value
                Q5 = Q5 & ", " & DistrictName
            End If
        End If
    Next row
End Function

