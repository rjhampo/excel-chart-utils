Attribute VB_Name = "Module1"
Dim ColorCodes As Object
Dim CHARTCONST() As Variant

Sub CellColorsToChart()
Attribute CellColorsToChart.VB_ProcData.VB_Invoke_Func = "E\n14"
    'To improve:
    '- Now does not need the scripting runtime to run (done)
    '- Combine specific and mass coloring (done)
    '- Fixed bar charts not coloring interior (done)
    '- Better customization options than just color (done)
    '- Need to implement string management for uniform inputs
    '- Script to apply to all charts, not just line and clustered bar

    Dim MultiColor As Boolean
    Dim CellMode As Boolean
    Dim xColor As Long
    Dim xCharts As ChartObjects
    Dim buf_series As Variant
    
    If TypeOf Selection Is Range Then
        For Each xCell In Selection
            If xCell.Value <> "" Then
                CellMode = True
                Exit For
            Else
                CellMode = False
            End If
        Next xCell
    End If
    If ActiveChart Is Nothing Then MultiColor = True Else MultiColor = False
    
    If CellMode Then
        For Each xCell In Selection
            xCell.Interior.Color = ColorCodes(UCase(xCell.Value2))(0)
            xCell.Font.Color = ColorCodes(UCase(xCell.Value2))(0)
        Next xCell
    Else
        If MultiColor Then
            Set xCharts = ActiveSheet.ChartObjects
            'If xCharts Is Nothing Then Exit Sub
            For Each xChart In xCharts
                For Each ser In xChart.Chart.SeriesCollection
                    buf_ser = UCase(ser.Name)
                    If IsInArray(ser.ChartType, CHARTCONST) Then
                        ser.Format.Fill.ForeColor.RGB = ColorCodes(buf_ser)(0)
                        If ColorCodes(buf_ser)(1) <> 0 Then
                            ser.Format.Fill.Patterned ColorCodes(buf_ser)(1)
                            ser.Format.Fill.BackColor.RGB = ColorCodes(buf_ser)(2)
                        End If
                    Else
                        ser.Format.Line.ForeColor.RGB = ColorCodes(buf_ser)(0)
                        ser.Format.Line.Weight = ColorCodes(buf_ser)(3)
                        ser.Format.Line.Transparency = ColorCodes(buf_ser)(11)
                        If ColorCodes(buf_ser)(4) = 1 Then
                            ser.Format.Line.DashStyle = ColorCodes(buf_ser)(5)
                        Else
                            ser.Format.Line.DashStyle = msoLineSolid
                        End If
                        If ColorCodes(buf_ser)(6) = 1 Then
                            ser.MarkerStyle = ColorCodes(buf_ser)(7)
                            ser.MarkerSize = ColorCodes(buf_ser)(8)
                            ser.MarkerForegroundColor = ColorCodes(buf_ser)(9)
                            ser.MarkerBackgroundColor = ColorCodes(buf_ser)(10)
                        Else
                            If ser.MarkerStyle <> -4142 Then
                               ser.MarkerStyle = -4142
                            End If
                        End If
                    End If
                Next ser
            Next xChart
        Else
            For Each ser In ActiveChart.SeriesCollection
                buf_ser = UCase(ser.Name)
                If IsInArray(ser.ChartType, CHARTCONST) Then
                    ser.Format.Fill.ForeColor.RGB = ColorCodes(buf_ser)(0)
                    If ColorCodes(buf_ser)(1) <> 0 Then
                        ser.Format.Fill.Patterned ColorCodes(buf_ser)(1)
                        ser.Format.Fill.BackColor.RGB = ColorCodes(buf_ser)(2)
                    End If
                Else
                    ser.Format.Line.ForeColor.RGB = ColorCodes(buf_ser)(0)
                    ser.Format.Line.Weight = ColorCodes(buf_ser)(3)
                    ser.Format.Line.Transparency = ColorCodes(buf_ser)(11)
                    If ColorCodes(buf_ser)(4) = 1 Then
                        ser.Format.Line.DashStyle = ColorCodes(buf_ser)(5)
                    Else
                        ser.Format.Line.DashStyle = msoLineSolid
                    End If
                    If ColorCodes(buf_ser)(6) = 1 Then
                        ser.MarkerStyle = ColorCodes(buf_ser)(7)
                        ser.MarkerSize = ColorCodes(buf_ser)(8)
                        ser.MarkerForegroundColor = ColorCodes(buf_ser)(9)
                        ser.MarkerBackgroundColor = ColorCodes(buf_ser)(10)
                    Else
                        If ser.MarkerStyle <> -4142 Then
                           ser.MarkerStyle = -4142
                        End If
                    End If
                End If
            Next ser
        End If
    End If
End Sub

Sub SetColorCodes()
Attribute SetColorCodes.VB_ProcData.VB_Invoke_Func = "W\n14"
    '- Join ResetColorCodes Sub to this Sub to make it concise (done)
    Set ColorCodes = CreateObject("Scripting.Dictionary")
    CHARTCONST = Array(57, 58, 59, 51, 52, 53) 'Filters bar and column 2d charts
    
    Dim xRg As Range
    Dim i As Integer
    Dim j As Integer
    Dim IsObtained As Boolean
    
    Dim DefaultColor As Long
    Dim BarPattern As Long
    Dim BarPatternBack As Long
    Dim Weight As Single
    Dim IsDashed As Integer
    Dim DashType As Long
    Dim IsMarker As Integer
    Dim MarkerType As Long
    Dim MarkerSize As Long
    Dim MarkerForeColor As Long
    Dim MarkerBackColor As Long
    Dim Transparency As Double
    
    IsObtained = False
    i = 1
    j = 0
    Set xRg = Selection
    
    For Each cell In xRg
        If Not IsObtained Then
            If xRg(i, 2).Interior.ColorIndex = xlNone Then
                DefaultColor = xlNone
            Else
                DefaultColor = xRg(i, 2).Interior.Color
            End If
            
            BarPattern = xRg(i, 3).Value
            
            If xRg(i, 4).Interior.ColorIndex = xlNone Then
                BarPatternBack = xlNone
            Else
                BarPatternBack = xRg(i, 4).Interior.Color
            End If
            
            If xRg(i, 5).Value = 0 Then
                Weight = 2.25
            Else
                Weight = xRg(i, 5).Value
            End If
            
            IsDashed = xRg(i, 6).Value
            DashType = xRg(i, 7).Value
            IsMarker = xRg(i, 8).Value
            MarkerType = xRg(i, 9).Value
            
            If xRg(i, 10).Value = 0 Then
                MarkerSize = 5
            Else
                MarkerSize = xRg(i, 10).Value
            End If
            
            If xRg(i, 11).Interior.ColorIndex = xlNone Then
                MarkerForeColor = xRg(i, 2).Interior.Color
            Else
                MarkerForeColor = xRg(i, 11).Interior.Color
            End If
            
            If xRg(i, 12).Interior.ColorIndex = xlNone Then
                MarkerBackColor = xRg(i, 2).Interior.Color
            Else
                MarkerBackColor = xRg(i, 12).Interior.Color
            End If
            
            Transparency = xRg(i, 13).Value
            
            ColorCodes.Add UCase(cell.Value), Array(DefaultColor, BarPattern, BarPatternBack, Weight, IsDashed, DashType, IsMarker, MarkerType, MarkerSize, MarkerForeColor, MarkerBackColor, Transparency)
            i = i + 1
            IsObtained = True
        Else
            j = j + 1
            If j >= 12 Then
                IsObtained = False
                j = 0
            End If
        End If
    Next cell
End Sub

Public Function IsInArray(valToBeFound, arr) As Boolean
    Dim i
    For i = LBound(arr) To UBound(arr)
        If arr(i) = valToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False

End Function
