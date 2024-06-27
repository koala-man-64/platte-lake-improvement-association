Attribute VB_Name = "Module4"
Sub Macro4()
' Macro4 Macro      Find days > 8
Dim xPoint(50) As Double, yPoint(50) As Double, Slope(50) As Double
Dim PointCount As Integer
Application.ScreenUpdating = False
Sheets("Lake Chemistry").Range("bt2").Select     'count number of fixed values
PointCount = ActiveCell.Value
ActiveCell.Offset(1, -2).Select
For i = 1 To PointCount + 1               'read x and y fixed values
                                        'and calculate slopes for each interval
    xPoint(i) = ActiveCell.Value
    ActiveCell.Offset(0, 1).Select
    yPoint(i) = ActiveCell.Value
    ActiveCell.Offset(1, -1).Select
Next i
For i = 1 To PointCount
    Slope(i) = (yPoint(i + 1) - yPoint(i)) / (xPoint(i + 1) - xPoint(i))
Next i
yint = yPoint(1)
If yint > 8 Then
    LowDOCount = 1
  Else
    LowDOCount = 0
End If
'Sheets("Sheet2").Range("L4").Select
For j = 1 To PointCount + 1                         'find next yInt and print
    For kk = xPoint(j) To xPoint(j + 1) - 1
      If kk > xPoint(j) And kk < xPoint(j + 1) Then SelectedSlope = Slope(j)
      yint = yint + SelectedSlope
      If yint > 8 Then LowDOCount = LowDOCount + 1
      'ActiveCell.Value = yint
      'ActiveCell.Offset(1, 0).Select
    Next kk
Next j
Sheets("Lake Chemistry").Range("bt3").Select
ActiveCell.Value = LowDOCount
End Sub