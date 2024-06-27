Attribute VB_Name = "Sheet16"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Attribute VB_Control = "CommandButton4, 5, 1, MSForms, CommandButton"
Attribute VB_Control = "CommandButton3, 4, 2, MSForms, CommandButton"
Attribute VB_Control = "CommandButton2, 3, 3, MSForms, CommandButton"
Attribute VB_Control = "CommandButton1, 1, 4, MSForms, CommandButton"
Attribute VB_Control = "TextBox1, 20, 5, MSForms, TextBox"
Private Sub CommandButton1_Click()
Dim Parameter As String, Site As String, Site1 As String, Site2 As String, Site3 As String
Dim TotalCount As Integer
Dim TestDate(5000) As Date, TestValue(5000) As Variant
Dim PlotDate(1000) As Date, PlotValue(1000) As Double
Dim SelectedYear As Integer
Dim i As Integer, j As Integer
Dim YAxisLabel As String, YColumn As String
Application.ScreenUpdating = False
Sheets("Stream Probe").Range("ba39:ba10000").ClearContents
Sheets("Stream Probe").Range("bc39:bc10000").ClearContents
Sheets("Stream Probe").Range("be39:be10000").ClearContents
Sheets("Stream Probe").Range("bg39:bg10000").ClearContents
Sheets("Stream Probe").Range("bi39:bi10000").ClearContents
Sheets("Stream Probe").Range("bk39:bk10000").ClearContents
Sheets("Stream Probe").Select
Sheets("Stream Probe").Range("j3").Select  'get Year1
SelectedYear = ActiveCell.Value
ActiveCell.Offset(1, 0).Select
Parameter = ActiveCell.Value        'get parameter
ActiveCell.Offset(1, 0).Select
Site1 = ActiveCell.Value      'get sites
ActiveCell.Offset(1, 0).Select
Site2 = ActiveCell.Value
ActiveCell.Offset(1, 0).Select
Site3 = ActiveCell.Value
SelectColumn = 2
ActiveSheet.Cells(35, SelectColumn).Select: MinYear = ActiveCell.Value  'min
ActiveSheet.Cells(36, SelectColumn).Select: MaxYear = ActiveCell.Value  'max
If SelectedYear > MaxYear Or SelectedYear < MinYear Then
    response% = MsgBox("Data are not available for Year = " & CStr(SelectedYear) & " as entered in Cell J3.", 64)
    response% = MsgBox("Enter Year between " & CStr(MinYear) & " and " & CStr(MaxYear) & ".", 64)
    Exit Sub
End If
Sheets("Stream Probe").Range("bb4").Select           'define chart 4 labels
ActiveCell.Value = Parameter + "    " + CStr(SelectedYear)
If Parameter = "Oxygen" Then YAxisLabel = "mg/L"
If Parameter = "Temperature" Then YAxisLabel = "degrees F"
If Parameter = "pH" Then YAxisLabel = "pH"
If Parameter = "Conductivity" Then YAxisLabel = "S/cm"
Sheets("Stream Probe").Range("bb5").Select: ActiveCell.Value = YAxisLabel
ActiveSheet.ChartObjects("Chart 4").Activate            'define chart 4 vertical scale
ActiveChart.ChartArea.Select
If Parameter = "Temperature" Then
    With ActiveChart.Axes(xlValue)
        .MinimumScale = 20
        .MaximumScale = 90
        .MajorUnit = 10
    End With
End If
If Parameter = "Oxygen" Then
    With ActiveChart.Axes(xlValue)
        .MinimumScale = 0
        .MaximumScale = 20
        .MajorUnit = 4
    End With
End If
If Parameter = "pH" Then
    With ActiveChart.Axes(xlValue)
        .MinimumScale = 5
        .MaximumScale = 10
        .MajorUnit = 1
    End With
End If
If Parameter = "Conductivity" Then
    With ActiveChart.Axes(xlValue)
        .MinimumScale = 0
        .MaximumScale = 500
        .MajorUnit = 100
    End With
End If
'***********************************************
Sheets("Stream Probe").Range("b37").Select
DataCount = ActiveCell.Value
ActiveCell.Offset(2, 0).Select           'get all the measure dates
For j = 1 To DataCount
      TestDate(j) = ActiveCell.Value     'select the first date record
      ActiveCell.Offset(1, 0).Select
Next j
For k = 1 To 3
    If k = 1 Then Site = Site1
    If k = 2 Then Site = Site2
    If k = 3 Then Site = Site3
    If Site = "Stone" And Parameter = "Oxygen" Then Column = "c"
    If Site = "Stone" And Parameter = "Temperature" Then Column = "d"
    If Site = "Stone" And Parameter = "pH" Then Column = "e"
    If Site = "Stone" And Parameter = "Conductivity" Then Column = "f"
    If Site = "Vet's" And Parameter = "Oxygen" Then Column = "h"
    If Site = "Vet's" And Parameter = "Temperature" Then Column = "i"
    If Site = "Vet's" And Parameter = "pH" Then Column = "j"
    If Site = "Vet's" And Parameter = "Conductivity" Then Column = "k"
    If Site = "Haze" And Parameter = "Oxygen" Then Column = "m"
    If Site = "Haze" And Parameter = "Temperature" Then Column = "n"
    If Site = "Haze" And Parameter = "pH" Then Column = "o"
    If Site = "Haze" And Parameter = "Conductivity" Then Column = "p"
    If Site = "Carter" And Parameter = "Oxygen" Then Column = "r"
    If Site = "Carter" And Parameter = "Temperature" Then Column = "s"
    If Site = "Carter" And Parameter = "pH" Then Column = "t"
    If Site = "Carter" And Parameter = "Conductivity" Then Column = "u"
    If Site = "Pioneer" And Parameter = "Oxygen" Then Column = "w"
    If Site = "Pioneer" And Parameter = "Temperature" Then Column = "x"
    If Site = "Pioneer" And Parameter = "pH" Then Column = "y"
    If Site = "Pioneer" And Parameter = "Conductivity" Then Column = "z"
    If Site = "USGS" And Parameter = "Oxygen" Then Column = "ab"
    If Site = "USGS" And Parameter = "Temperature" Then Column = "ac"
    If Site = "USGS" And Parameter = "pH" Then Column = "ad"
    If Site = "USGS" And Parameter = "Conductivity" Then Column = "ae"
    If Site = "NB Ind Hill" And Parameter = "Oxygen" Then Column = "ag"
    If Site = "NB Ind Hill" And Parameter = "Temperature" Then Column = "ah"
    If Site = "NB Ind Hill" And Parameter = "pH" Then Column = "ai"
    If Site = "NB Ind Hill" And Parameter = "Conductivity" Then Column = "aj"
    If Site = "NB Dead" And Parameter = "Oxygen" Then Column = "al"
    If Site = "NB Dead" And Parameter = "Temperature" Then Column = "am"
    If Site = "NB Dead" And Parameter = "pH" Then Column = "an"
    If Site = "NB Dead" And Parameter = "Conductivity" Then Column = "ao"
    kk = 1
    Sheets("Stream Probe").Range(Column + CStr(39)).Select
    For j = 1 To DataCount
      TestValue(j) = ActiveCell.Value     'select the first date record
      ActiveCell.Offset(1, 0).Select
    Next j
    For i = 1 To DataCount
            If Year(TestDate(i)) > SelectedYear Then Exit For
            If Year(TestDate(i)) = SelectedYear Then
                PlotDate(kk) = TestDate(i)
                PlotValue(kk) = TestValue(i)
                kk = kk + 1                   'increment the count
            End If
            ActiveCell.Offset(1, 0).Select     'go down one and back to the date column
     Next i
    'Chart Data
    If k = 1 Then Sheets("Stream Probe").Range("ba39").Select
    If k = 2 Then Sheets("Stream Probe").Range("be39").Select
    If k = 3 Then Sheets("Stream Probe").Range("bi39").Select
    For j = 1 To kk - 1
            ActiveCell.Value = PlotDate(j)
            ActiveCell.Offset(0, 2).Select
            If PlotValue(j) <> 0 Then ActiveCell.Value = PlotValue(j)
            ActiveCell.Offset(1, -2).Select
    Next j
Next k
Sheets("Stream Probe").Range("bb14").Select
ActiveCell.Value = Site1
ActiveCell.Offset(1, 0).Select
ActiveCell.Value = Site2
ActiveCell.Offset(1, 0).Select
ActiveCell.Value = Site3
Sheets("Stream Probe").Range("j3").Select
End Sub
Private Sub CommandButton2_Click()
Dim Parameter As String, Site As String, Site1 As String, Site2 As String, Site3 As String
Dim TotalCount As Integer
Dim TestDate(5000) As Date, TestValue(5000) As Variant
Dim PlotDate(1000) As Date, PlotValue(1000) As Double
Dim SelectedYear As Integer
Dim i As Integer, j As Integer
Dim YAxisLabel As String, YColumn As String
Application.ScreenUpdating = False
Sheets("Stream Probe").Range("bm39:bm10000").ClearContents
Sheets("Stream Probe").Range("bo39:bo10000").ClearContents
Sheets("Stream Probe").Range("bq39:bq10000").ClearContents
Sheets("Stream Probe").Range("bs39:bs10000").ClearContents
Sheets("Stream Probe").Range("bu39:bu10000").ClearContents
Sheets("Stream Probe").Range("bw39:bw10000").ClearContents
Sheets("Stream Probe").Select
Sheets("Stream Probe").Range("n3").Select  'get Year1
SelectedYear = ActiveCell.Value
ActiveCell.Offset(1, 0).Select
Parameter = ActiveCell.Value        'get parameter
ActiveCell.Offset(1, 0).Select
Site1 = ActiveCell.Value      'get sites
ActiveCell.Offset(1, 0).Select
Site2 = ActiveCell.Value
ActiveCell.Offset(1, 0).Select
Site3 = ActiveCell.Value
SelectColumn = 2
ActiveSheet.Cells(35, SelectColumn).Select: MinYear = ActiveCell.Value  'min
ActiveSheet.Cells(36, SelectColumn).Select: MaxYear = ActiveCell.Value  'max
If SelectedYear > MaxYear Or SelectedYear < MinYear Then
    response% = MsgBox("Data are not available for Year = " & CStr(SelectedYear) & " as entered in Cell N3.", 64)
    response% = MsgBox("Enter Year between " & CStr(MinYear) & " and " & CStr(MaxYear) & ".", 64)
    Exit Sub
End If
Sheets("Stream Probe").Range("bb9").Select           'define chart 4 labels
ActiveCell.Value = Parameter + "    " + CStr(SelectedYear)
If Parameter = "Oxygen" Then YAxisLabel = "mg/L"
If Parameter = "Temperature" Then YAxisLabel = "degrees F"
If Parameter = "pH" Then YAxisLabel = "pH"
If Parameter = "Conductivity" Then YAxisLabel = "S/cm"
Sheets("Stream Probe").Range("bb10").Select: ActiveCell.Value = YAxisLabel
ActiveSheet.ChartObjects("Chart 5").Activate            'define chart 4 vertical scale
ActiveChart.ChartArea.Select
If Parameter = "Temperature" Then
    With ActiveChart.Axes(xlValue)
        .MinimumScale = 20
        .MaximumScale = 90
        .MajorUnit = 10
    End With
End If
If Parameter = "Oxygen" Then
    With ActiveChart.Axes(xlValue)
        .MinimumScale = 0
        .MaximumScale = 20
        .MajorUnit = 4
    End With
End If
If Parameter = "pH" Then
    With ActiveChart.Axes(xlValue)
        .MinimumScale = 5
        .MaximumScale = 10
        .MajorUnit = 1
    End With
End If
If Parameter = "Conductivity" Then
    With ActiveChart.Axes(xlValue)
        .MinimumScale = 0
        .MaximumScale = 500
        .MajorUnit = 100
    End With
End If
'***********************************************
Sheets("Stream Probe").Range("b37").Select
DataCount = ActiveCell.Value
ActiveCell.Offset(2, 0).Select           'get all the measure dates
For j = 1 To DataCount
      TestDate(j) = ActiveCell.Value     'select the first date record
      ActiveCell.Offset(1, 0).Select
Next j
For k = 1 To 3
    If k = 1 Then Site = Site1
    If k = 2 Then Site = Site2
    If k = 3 Then Site = Site3
    If Site = "Stone" And Parameter = "Oxygen" Then Column = "c"
    If Site = "Stone" And Parameter = "Temperature" Then Column = "d"
    If Site = "Stone" And Parameter = "pH" Then Column = "e"
    If Site = "Stone" And Parameter = "Conductivity" Then Column = "f"
    If Site = "Vet's" And Parameter = "Oxygen" Then Column = "h"
    If Site = "Vet's" And Parameter = "Temperature" Then Column = "i"
    If Site = "Vet's" And Parameter = "pH" Then Column = "j"
    If Site = "Vet's" And Parameter = "Conductivity" Then Column = "k"
    If Site = "Haze" And Parameter = "Oxygen" Then Column = "m"
    If Site = "Haze" And Parameter = "Temperature" Then Column = "n"
    If Site = "Haze" And Parameter = "pH" Then Column = "o"
    If Site = "Haze" And Parameter = "Conductivity" Then Column = "p"
    If Site = "Carter" And Parameter = "Oxygen" Then Column = "r"
    If Site = "Carter" And Parameter = "Temperature" Then Column = "s"
    If Site = "Carter" And Parameter = "pH" Then Column = "t"
    If Site = "Carter" And Parameter = "Conductivity" Then Column = "u"
    If Site = "Pioneer" And Parameter = "Oxygen" Then Column = "w"
    If Site = "Pioneer" And Parameter = "Temperature" Then Column = "x"
    If Site = "Pioneer" And Parameter = "pH" Then Column = "y"
    If Site = "Pioneer" And Parameter = "Conductivity" Then Column = "z"
    If Site = "USGS" And Parameter = "Oxygen" Then Column = "ab"
    If Site = "USGS" And Parameter = "Temperature" Then Column = "ac"
    If Site = "USGS" And Parameter = "pH" Then Column = "ad"
    If Site = "USGS" And Parameter = "Conductivity" Then Column = "ae"
    If Site = "NB Ind Hill" And Parameter = "Oxygen" Then Column = "ag"
    If Site = "NB Ind Hill" And Parameter = "Temperature" Then Column = "ah"
    If Site = "NB Ind Hill" And Parameter = "pH" Then Column = "ai"
    If Site = "NB Ind Hill" And Parameter = "Conductivity" Then Column = "aj"
    If Site = "NB Dead" And Parameter = "Oxygen" Then Column = "al"
    If Site = "NB Dead" And Parameter = "Temperature" Then Column = "am"
    If Site = "NB Dead" And Parameter = "pH" Then Column = "an"
    If Site = "NB Dead" And Parameter = "Conductivity" Then Column = "ao"
    kk = 1
    Sheets("Stream Probe").Range(Column + CStr(39)).Select
    For j = 1 To DataCount
      TestValue(j) = ActiveCell.Value     'select the first date record
      ActiveCell.Offset(1, 0).Select
    Next j
    For i = 1 To DataCount
            If Year(TestDate(i)) > SelectedYear Then Exit For
            If Year(TestDate(i)) = SelectedYear Then
                PlotDate(kk) = TestDate(i)
                PlotValue(kk) = TestValue(i)
                kk = kk + 1                   'increment the count
            End If
            ActiveCell.Offset(1, 0).Select     'go down one and back to the date column
     Next i
    'Chart Data
    If k = 1 Then Sheets("Stream Probe").Range("bm39").Select
    If k = 2 Then Sheets("Stream Probe").Range("bq39").Select
    If k = 3 Then Sheets("Stream Probe").Range("bu39").Select
    For j = 1 To kk - 1
            ActiveCell.Value = PlotDate(j)
            ActiveCell.Offset(0, 2).Select
            If PlotValue(j) <> 0 Then ActiveCell.Value = PlotValue(j)
            ActiveCell.Offset(1, -2).Select
    Next j
Next k
Sheets("Stream Probe").Range("be14").Select
ActiveCell.Value = Site1
ActiveCell.Offset(1, 0).Select
ActiveCell.Value = Site2
ActiveCell.Offset(1, 0).Select
ActiveCell.Value = Site3
Sheets("Stream Probe").Range("n3").Select
End Sub
Private Sub CommandButton3_Click()
    Sheets("Main Menu").Select
    Sheets("Main Menu").Range("g11").Select
End Sub
Private Sub CommandButton4_Click()
If CommandButton4.Caption = "Open" Then
        CommandButton4.Caption = "Close"
        TextBox1.Visible = True
    Else
        CommandButton4.Caption = "Open"
        TextBox1.Visible = False
    End If
End Sub
Private Sub worksheet_Activate()
Worksheets("Stream Probe").Activate
    Sheets("Stream Probe").Select
    TextBox1.Visible = False
    CommandButton4.Caption = "Open"
End Sub