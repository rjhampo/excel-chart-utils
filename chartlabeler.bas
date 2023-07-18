Attribute VB_Name = "Module3"
Dim labelData() As String

Sub Labelizer()
Attribute Labelizer.VB_ProcData.VB_Invoke_Func = "X\n14"
    ' To improve:
    ' - Make it label the series as formulas instead of values
    
    Dim xChart As Chart
    Dim numRow As Long
    Dim numCol As Long
    Dim cellAdd As String
    Set xChart = ActiveChart
    numRow = 1
    numCol = 1
    For Each xSeriesCol In xChart.FullSeriesCollection
        For Each xPoint In xSeriesCol.Points
            xPoint.DataLabel.Formula = "='" & ActiveSheet.Name & "'!" & labelData(numRow, numCol)
            numCol = numCol + 1
        Next xPoint
        numCol = 1
        numRow = numRow + 1
    Next xSeriesCol
End Sub

Sub SeedLabelizer()
Attribute SeedLabelizer.VB_ProcData.VB_Invoke_Func = "N\n14"
    Dim numRow As Long
    Dim numCol As Long
    Dim xSelection As Range
    Set xSelection = Selection
    numRow = xSelection.Rows.Count
    numCol = xSelection.Columns.Count
    ReDim labelData(numRow, numCol) As String
    
    For i = 1 To numRow
        For j = 1 To numCol
            labelData(i, j) = xSelection(i, j).Address
        Next j
    Next i
End Sub
