Attribute VB_Name = "Sheet1"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Attribute VB_Control = "CommandButton5, 879, 0, MSForms, CommandButton"
Attribute VB_Control = "CommandButton1, 1, 1, MSForms, CommandButton"
Attribute VB_Control = "CommandButton2, 10, 2, MSForms, CommandButton"
Attribute VB_Control = "CommandButton3, 77, 3, MSForms, CommandButton"
Attribute VB_Control = "CommandButton4, 875, 4, MSForms, CommandButton"
Attribute VB_Control = "TextBox1, 876, 5, MSForms, TextBox"
Private Sub CommandButton1_Click()
Dim TestDate As Date, TestValue As Variant
Dim PlotDate(50000) As Date, PlotValue(50000) As Variant
Dim TotalTestDays As Integer, PointCount As Integer
Dim xPoint(100) As Double, yPoint(100) As Double, Slope(100) As Double
Dim LowDOCount(5) As Double
Dim AddTOListDate(5000) As Date, AddToListValue(5000) As Double
ActiveSheet.ChartObjects("Chart 14").Activate
ActiveSheet.Shapes("Chart 14").ZOrder msoBringToFront
ActiveSheet.ChartObjects("Chart 17").Activate
ActiveSheet.Shapes("Chart 17").ZOrder msoBringToFront
Sheets("Lake Probe Data").Select
Sheets("Lake Probe Data").Range("c37").Select: DataCount = ActiveCell.Value
Sheets("Lake Probe Data").Range("h3").Select: SelectedYear = ActiveCell.Value
Sheets("Lake Probe Data").Range("h4").Select: Parameter = ActiveCell.Value
Sheets("Lake Probe Data").Range("aj13").Select:  ActiveCell.Value = Parameter
SelectColumn = 2
ActiveSheet.Cells(37, SelectColumn).Select: MinYear = ActiveCell.Value  'min
ActiveSheet.Cells(38, SelectColumn).Select: MaxYear = ActiveCell.Value  'max
If SelectedYear > MaxYear Or SelectedYear < MinYear Then
    response% = MsgBox("Data are not available for Year = " & CStr(SelectedYear) & " as entered in Cell H3.", 64)
    response% = MsgBox("Enter Year between " & CStr(MinYear) & " and " & CStr(MaxYear) & ".", 64)
    Exit Sub
End If
Application.ScreenUpdating = False
If Parameter = "Oxygen" Then YAxisLabel = "mg/L"
If Parameter = "Temperature" Then YAxisLabel = "degrees F"
If Parameter = "pH" Then YAxisLabel = "pH"
If Parameter = "ORP" Then YAxisLabel = "mV"
If Parameter = "Conductivity" Then YAxisLabel = "S/cm"
Sheets("Lake Probe Data").Range("ae5").Select: ActiveCell.Value = YAxisLabel
ActiveSheet.ChartObjects("Chart 14").Activate
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
            .MaximumScale = 15
            .MajorUnit = 3
        End With
End If
If Parameter = "pH" Then
        With ActiveChart.Axes(xlValue)
            .MinimumScale = 5
            .MaximumScale = 10
            .MajorUnit = 1
        End With
End If
If Parameter = "ORP" Then
        With ActiveChart.Axes(xlValue)
            .MinimumScale = -300
            .MaximumScale = 400
            .MajorUnit = 100
        End With
End If
If Parameter = "Conductivity" Then
    ActiveChart.Axes(xlValue).Select
    Selection.TickLabels.NumberFormat = "0"
        With ActiveChart.Axes(xlValue)
            .MinimumScale = 0
            .MaximumScale = 600
            .MajorUnit = 100
        End With
End If
Sheets("Lake Probe Data").Range("k43:k93").ClearContents 'clear plot data from previous run
Sheets("Lake Probe Data").Range("m43:t93").ClearContents 'clear plot data from previous run
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sheets("Lake Probe Data").Range("b39").Select        'first date
k = 1                       'number of dates during SelectedYear
For j = 1 To DataCount      'scan through all the data
    TestDate = ActiveCell.Value
    If Year(TestDate) > SelectedYear Then Exit For
    If Year(TestDate) = SelectedYear Then
        PlotDate(k) = TestDate
        If Parameter = "Temperature" Then ActiveCell.Offset(0, 2).Select     'move to right to find the correct record
        If Parameter = "Oxygen" Then ActiveCell.Offset(0, 3).Select
        If Parameter = "ORP" Then ActiveCell.Offset(0, 4).Select
        If Parameter = "Conductivity" Then ActiveCell.Offset(0, 5).Select
        If Parameter = "pH" Then ActiveCell.Offset(0, 6).Select
        TestValue = ActiveCell.Value    'select the first date parameter value
        PlotValue(k) = TestValue
        If Parameter = "Oxygen" Then ActiveCell.Offset(0, -3).Select     'go down one and back to the date column
        If Parameter = "Temperature" Then ActiveCell.Offset(0, -2).Select
        If Parameter = "ORP" Then ActiveCell.Offset(0, -4).Select
        If Parameter = "Conductivity" Then ActiveCell.Offset(0, -5).Select
        If Parameter = "pH" Then ActiveCell.Offset(0, -6).Select
        k = k + 1
     End If
     ActiveCell.Offset(1, 0).Select
Next j
TotalTestDays = k - 1                             'need to substract one because of k = k + 1 above
PointCount = TotalTestDays / 8
Sheets("Lake Probe Data").Range("k43").Select     'print dates for plot
For i = 1 To TotalTestDays Step 8
    ActiveCell.Value = PlotDate(i)
    ActiveCell.Offset(1, 0).Select
Next i
For kk = 1 To 8
    If kk = 1 Then Sheets("Lake Probe Data").Range("m43").Select   'print values for 8 depths for plot
    If kk = 2 Then Sheets("Lake Probe Data").Range("n43").Select
    If kk = 3 Then Sheets("Lake Probe Data").Range("o43").Select
    If kk = 4 Then Sheets("Lake Probe Data").Range("p43").Select
    If kk = 5 Then Sheets("Lake Probe Data").Range("q43").Select
    If kk = 6 Then Sheets("Lake Probe Data").Range("r43").Select
    If kk = 7 Then Sheets("Lake Probe Data").Range("s43").Select
    If kk = 8 Then Sheets("Lake Probe Data").Range("t43").Select
    For i = kk To TotalTestDays - 8 + kk Step 8
            ActiveCell.Value = PlotValue(i)
            ActiveCell.Offset(1, 0).Select
     Next i
 Next kk
'Calculate Number of Low DO days    'xxxxxxxxxxxxxxxxxxxxx
If Parameter = "Oxygen" Then
    For k = 1 To 4
        Sheets("Lake Probe Data").Range("L43").Select  'find the xPoint values = day of the year
        For i = 1 To PointCount
            xPoint(i) = ActiveCell.Value
            ActiveCell.Offset(1, 0).Select
        Next i
        If k = 1 Then Sheets("Lake Probe Data").Range("q43").Select  'find the yPoint values for 45 foot
        If k = 2 Then Sheets("Lake Probe Data").Range("r43").Select  'find the yPoint values for 60 foot
        If k = 3 Then Sheets("Lake Probe Data").Range("s43").Select  'find the yPoint values for 75 foot
        If k = 4 Then Sheets("Lake Probe Data").Range("t43").Select  'find the yPoint values for 90 foot
        For i = 1 To PointCount
            yPoint(i) = ActiveCell.Value
            ActiveCell.Offset(1, 0).Select
        Next i
        For i = 2 To PointCount           'calculate l3 slopes
            Slope(i - 1) = (yPoint(i) - yPoint(i - 1)) / (xPoint(i) - xPoint(i - 1))
        Next i
        LowDOCount(k) = 0                                                 'set first yInt and LowDOCount
        yint = yPoint(1)
         For DOY = xPoint(1) + 1 To xPoint(PointCount) 'DOY = 50 to 307  step 1                                   'DOY goes over whole range
            For kk = 1 To PointCount                                                        'kk is the panel
              If DOY > xPoint(kk) And DOY < xPoint(kk + 1) Then
                SelectedSlope = Slope(kk)  'this finds the correct slope for the panel
                Exit For
              End If
            Next kk
            yint = yint + SelectedSlope                                                     'this is OK because x(j)-x(j-1) is always 1
            If yint < 2 Then LowDOCount(k) = LowDOCount(k) + 1
         Next DOY
   Next k
   SedRelease = Round((LowDOCount(1) * 0.411 * 1149270 + LowDOCount(2) * 0.411 * 1023821 + LowDOCount(3) * 0.411 * 473467 + LowDOCount(4) * 1.547 * 105215) * 0.000002204, 1)
    ' 12   45   75      57.4
   Sheets("Annual Averages").Select
   Sheets("Annual Averages").Range("e" + CStr(SelectedYear - 2010 + 48)).Select
   ActiveCell.Value = SedRelease
   Sheets("Lake Probe Data").Select
End If              'xxxxxxxxxxxxxxxxxxxxxxxxxxxxx
If Parameter = "Oxygen" Then
    Sheets("Lake Probe Data").Range("ae4").Select: ActiveCell.Value = CStr(SelectedYear) + "  " + Parameter + "   " + "Sed Rel = " + CStr(SedRelease) 'define chart title
    Sheets("Lake Probe Data").Range("ai4").Select
    ActiveCell.Value = LowDOCount(1)
    ActiveCell.Offset(1, 0).Select
    ActiveCell.Value = LowDOCount(2)
    ActiveCell.Offset(1, 0).Select
    ActiveCell.Value = LowDOCount(3)
    ActiveCell.Offset(1, 0).Select
    ActiveCell.Value = LowDOCount(4)
End If
If Parameter = "Temperature" Then
    Sheets("Lake Probe Data").Range("m94").Select: MaxT = ActiveCell.Value
    Sheets("Lake Probe Data").Range("ae4").Select: ActiveCell.Value = CStr(SelectedYear) + "  " + Parameter + "   " + "Max = " + CStr(MaxT) 'define chart title
End If
If Parameter = "ORP" Then
    Sheets("Lake Probe Data").Range("n94").Select: MinORP = ActiveCell.Value
    Sheets("Lake Probe Data").Range("ae4").Select: ActiveCell.Value = CStr(SelectedYear) + "  " + Parameter + "   " + "Min = " + CStr(MinORP) 'define chart title
End If
If Parameter = "Conductivity" Then
    Sheets("Lake Probe Data").Range("o94").Select: AvgC = ActiveCell.Value
    AvgC = Round(AvgC, 1)
    Sheets("Lake Probe Data").Range("ae4").Select: ActiveCell.Value = CStr(SelectedYear) + "  " + Parameter + "   " + "Avg = " + CStr(AvgC) 'define chart title
End If
If Parameter = "pH" Then
    Sheets("Lake Probe Data").Range("p94").Select: MinpH = ActiveCell.Value
    Sheets("Lake Probe Data").Range("ae4").Select: ActiveCell.Value = CStr(SelectedYear) + "  " + Parameter + "   " + "Min = " + CStr(MinpH) 'define chart title
End If
'HHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHH
'HHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHH
'HHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHH
If Parameter = "Temperature" Then ''''''''''''''''''''''''''''''''''''''''''get the new VolwtValues
    Sheets("Lake Probe Data").Range("v93").Select   'find last value in col V'''''''''''''''
    For i = 1 To 51
       If ActiveCell.Value > 0 Then
         VolWtValue = ActiveCell.Value
         ActiveCell.Offset(0, -11).Select
         VolWtDate = ActiveCell.Value
         Exit For
       End If
       ActiveCell.Offset(-1, 0).Select
     Next i
     Sheets("Lake Probe Data").Range("x42").Select: MaxDate = ActiveCell.Value      '''''''''''''''''''''''
     Sheets("Lake Probe Data").Range("y42").Select: ListCount = ActiveCell.Value                                'start print sequence
     PrintRow = ListCount + 43
     Sheets("Lake Probe Data").Range("x" + CStr(PrintRow)).Select      'print List values
      If VolWtValue > 0 Then
        ActiveCell.Value = VolWtDate
        ActiveCell.Offset(0, 1).Select
        ActiveCell.Value = VolWtValue
      End If
End If
Sheets("Lake Probe Data").Range("r13").Select
End Sub
Private Sub CommandButton2_Click()
Dim TestDate As Date, TestValue As Variant
Dim PlotDate(50000) As Date, PlotValue(50000) As Variant
Dim TotalTestDays As Integer, PointCount As Integer
Dim xPoint(50) As Double, yPoint(50) As Double, Slope(50) As Double
Dim LowDOCount(5) As Double
'Lake Probe DATA  Chart 2
ActiveSheet.ChartObjects("Chart 14").Activate
ActiveSheet.Shapes("Chart 14").ZOrder msoBringToFront
ActiveSheet.ChartObjects("Chart 17").Activate
ActiveSheet.Shapes("Chart 17").ZOrder msoBringToFront
Application.ScreenUpdating = False
Sheets("Lake Probe Data").Select
Sheets("Lake Probe Data").Range("c37").Select: DataCount = ActiveCell.Value
Sheets("Lake Probe Data").Range("l3").Select: SelectedYear = ActiveCell.Value
Sheets("Lake Probe Data").Range("l4").Select: Parameter = ActiveCell.Value
Sheets("Lake Probe Data").Range("aj14").Select: ActiveCell.Value = Parameter
SelectColumn = 2
ActiveSheet.Cells(37, SelectColumn).Select: MinYear = ActiveCell.Value  'min
ActiveSheet.Cells(38, SelectColumn).Select: MaxYear = ActiveCell.Value  'max
If SelectedYear > MaxYear Or SelectedYear < MinYear Then
    response% = MsgBox("Data are not available for Year = " & CStr(SelectedYear) & " as entered in Cell L3.", 64)
    response% = MsgBox("Enter Year between " & CStr(MinYear) & " and " & CStr(MaxYear) & ".", 64)
    Exit Sub
End If
If Parameter = "Oxygen" Then YAxisLabel = "mg/L"
If Parameter = "Temperature" Then YAxisLabel = "degrees F"
If Parameter = "pH" Then YAxisLabel = "pH"
If Parameter = "ORP" Then YAxisLabel = "mV"
If Parameter = "Conductivity" Then YAxisLabel = "S/cm"
Sheets("Lake Probe Data").Range("ae10").Select: ActiveCell.Value = YAxisLabel
ActiveSheet.ChartObjects("Chart 17").Activate
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
            .MaximumScale = 15
            .MajorUnit = 3
        End With
End If
If Parameter = "pH" Then
        With ActiveChart.Axes(xlValue)
            .MinimumScale = 5
            .MaximumScale = 10
            .MajorUnit = 1
        End With
End If
If Parameter = "ORP" Then
        With ActiveChart.Axes(xlValue)
            .MinimumScale = -300
            .MaximumScale = 400
            .MajorUnit = 100
        End With
End If
If Parameter = "Conductivity" Then
    ActiveChart.Axes(xlValue).Select
    Selection.TickLabels.NumberFormat = "0"
        With ActiveChart.Axes(xlValue)
            .MinimumScale = 0
            .MaximumScale = 600
            .MajorUnit = 100
        End With
End If
Sheets("Lake Probe Data").Range("k100:k150").ClearContents 'clear plot data from previous run
Sheets("Lake Probe Data").Range("m100:t150").ClearContents 'clear plot data from previous run
Sheets("Lake Probe Data").Range("b39").Select        'first date
k = 1
For j = 1 To DataCount      'scan through all the data
    TestDate = ActiveCell.Value
    If Year(TestDate) > SelectedYear Then Exit For
    If Year(TestDate) = SelectedYear Then
        PlotDate(k) = TestDate
        If Parameter = "Temperature" Then ActiveCell.Offset(0, 2).Select     'move to right to find the correct record
        If Parameter = "Oxygen" Then ActiveCell.Offset(0, 3).Select
        If Parameter = "ORP" Then ActiveCell.Offset(0, 4).Select
        If Parameter = "Conductivity" Then ActiveCell.Offset(0, 5).Select
        If Parameter = "pH" Then ActiveCell.Offset(0, 6).Select
        TestValue = ActiveCell.Value    'select the first date parameter value
        PlotValue(k) = TestValue
        If Parameter = "Oxygen" Then ActiveCell.Offset(0, -3).Select     'go down one and back to the date column
        If Parameter = "Temperature" Then ActiveCell.Offset(0, -2).Select
        If Parameter = "ORP" Then ActiveCell.Offset(0, -4).Select
        If Parameter = "Conductivity" Then ActiveCell.Offset(0, -5).Select
        If Parameter = "pH" Then ActiveCell.Offset(0, -6).Select
        k = k + 1
     End If
     ActiveCell.Offset(1, 0).Select
Next j
TotalTestDays = k - 1               'need to substract one because of k = k + 1 above
PointCount = TotalTestDays / 8
Sheets("Lake Probe Data").Range("k100").Select     'print dates for plot
For i = 1 To TotalTestDays Step 8
    ActiveCell.Value = PlotDate(i)
    ActiveCell.Offset(1, 0).Select
Next i
For kk = 1 To 8
    If kk = 1 Then Sheets("Lake Probe Data").Range("m100").Select   'print values for 8 depths for plot
    If kk = 2 Then Sheets("Lake Probe Data").Range("n100").Select
    If kk = 3 Then Sheets("Lake Probe Data").Range("o100").Select
    If kk = 4 Then Sheets("Lake Probe Data").Range("p100").Select
    If kk = 5 Then Sheets("Lake Probe Data").Range("q100").Select
    If kk = 6 Then Sheets("Lake Probe Data").Range("r100").Select
    If kk = 7 Then Sheets("Lake Probe Data").Range("s100").Select
    If kk = 8 Then Sheets("Lake Probe Data").Range("t100").Select
    For i = kk To TotalTestDays - 8 + kk Step 8
        ActiveCell.Value = PlotValue(i)
        ActiveCell.Offset(1, 0).Select
    Next i
 Next kk
 'Calculate Number of Low DO days       xxxxxxxxxxxxxxxxxxxxx
If Parameter = "Oxygen" Then        'xxxxxxxxxxxxxxxxxxxxx
    For k = 1 To 4
        Sheets("Lake Probe Data").Range("L100").Select  'find the xPoint values
        For i = 1 To PointCount
            xPoint(i) = ActiveCell.Value
            ActiveCell.Offset(1, 0).Select
        Next i
        If k = 1 Then Sheets("Lake Probe Data").Range("q100").Select 'find the yPoint values for 45 foot
        If k = 2 Then Sheets("Lake Probe Data").Range("r100").Select  'find the yPoint values for 60 foot
        If k = 3 Then Sheets("Lake Probe Data").Range("s100").Select  'find the yPoint values for 75 foot
        If k = 4 Then Sheets("Lake Probe Data").Range("t100").Select  'find the yPoint values for 90 foot
        For i = 1 To PointCount
            yPoint(i) = ActiveCell.Value
            ActiveCell.Offset(1, 0).Select
        Next i
        For i = 2 To PointCount            'calculate slopes
            Slope(i - 1) = (yPoint(i) - yPoint(i - 1)) / (xPoint(i) - xPoint(i - 1))        'artificial slope(0)not used
        Next i
        LowDOCount(k) = 0                                                                   'set first yInt and LowDOCount
        yint = yPoint(1)                                                                     'set first yInt and LowDOCount
        For DOY = xPoint(1) + 1 To xPoint(PointCount)                                       'DOY goes over whole range
            For kk = 1 To PointCount                                                        'kk is the panel
              If DOY > xPoint(kk) And DOY < xPoint(kk + 1) Then
                SelectedSlope = Slope(kk)  'this finds the correct slope for the panel
                Exit For
              End If
            Next kk
            yint = yint + SelectedSlope                                                     'this is OK because x(j)-x(j-1) is always 1
            If yint < 2 Then LowDOCount(k) = LowDOCount(k) + 1
        Next DOY
   Next k
   SedRelease = Round((LowDOCount(1) * 0.411 * 1149270 + LowDOCount(2) * 0.411 * 1023821 + LowDOCount(3) * 0.411 * 473467 + LowDOCount(4) * 1.547 * 105215) * 0.000002204, 1)
   Sheets("Annual Averages").Select
   Sheets("Annual Averages").Range("e" + CStr(SelectedYear - 2010 + 48)).Select
   ActiveCell.Value = SedRelease
   Sheets("Lake Probe Data").Select
End If              'xxxxxxxxxxxxxxxxxxxxxxxxxxxxx
If Parameter = "Oxygen" Then
    Sheets("Lake Probe Data").Range("ae9").Select: ActiveCell.Value = CStr(SelectedYear) + "  " + Parameter + "   " + "Sed Rel = " + CStr(SedRelease) 'define chart title
    Sheets("Lake Probe Data").Range("aj4").Select
    ActiveCell.Value = LowDOCount(1)
    ActiveCell.Offset(1, 0).Select
    ActiveCell.Value = LowDOCount(2)
    ActiveCell.Offset(1, 0).Select
    ActiveCell.Value = LowDOCount(3)
    ActiveCell.Offset(1, 0).Select
    ActiveCell.Value = LowDOCount(4)
End If
If Parameter = "Temperature" Then
    Sheets("Lake Probe Data").Range("m97").Select: MaxT = ActiveCell.Value
    Sheets("Lake Probe Data").Range("ae9").Select: ActiveCell.Value = CStr(SelectedYear) + "  " + Parameter + "   " + "Max = " + CStr(MaxT) 'define chart title
End If
If Parameter = "ORP" Then
    Sheets("Lake Probe Data").Range("n97").Select: MinORP = ActiveCell.Value
    Sheets("Lake Probe Data").Range("ae9").Select: ActiveCell.Value = CStr(SelectedYear) + "  " + Parameter + "   " + "Min = " + CStr(MinORP) 'define chart title
End If
If Parameter = "Conductivity" Then
    Sheets("Lake Probe Data").Range("o97").Select: AvgC = ActiveCell.Value
    AvgC = Round(AvgC, 1)    'round
    Sheets("Lake Probe Data").Range("ae9").Select: ActiveCell.Value = CStr(SelectedYear) + "  " + Parameter + "   " + "Avg = " + CStr(AvgC) 'define chart title
End If
If Parameter = "pH" Then
    Sheets("Lake Probe Data").Range("p97").Select: MinpH = ActiveCell.Value
    Sheets("Lake Probe Data").Range("ae9").Select: ActiveCell.Value = CStr(SelectedYear) + "  " + Parameter + "   " + "Min = " + CStr(MinpH) 'define chart title
End If
Sheets("Lake Probe Data").Range("L3").Select
End Sub
Private Sub CommandButton5_Click()
Sheets("Lake Probe Data").Range("aj13").Select
If ActiveCell.Value <> "Temperature" Then            'check that data are available for selected range
    response% = MsgBox("The Chart 1 parameter must Temperature.  The Chart 2 parameter must be Oxygen.  The Chart 1 Year and the Chart 2 Year must be the same.", 64)
    ActiveWindow.ScrollColumn = 1
    ActiveWindow.ScrollRow = 1
    Sheets("Lake Probe Data").Range("h3").Select
    Exit Sub
End If
Sheets("Lake Probe Data").Range("aj14").Select
If ActiveCell.Value <> "Oxygen" Then            'check that data are available for selected range
    response% = MsgBox("The Chart 1 parameter must Temperature.  The Chart 2 parameter must be Oxygen.  The Chart 1 Year and the Chart 2 Year must be the same.", 64)
    ActiveWindow.ScrollColumn = 1
    ActiveWindow.ScrollRow = 1
    Sheets("Lake Probe Data").Range("h3").Select
    Exit Sub
End If
Sheets("Lake Probe Data").Range("k43").Select: Yr1 = Year(ActiveCell.Value)
Sheets("Lake Probe Data").Range("K100").Select: Yr2 = Year(ActiveCell.Value)
If Yr1 <> Yr2 Then            'check that data are available for selected range
    response% = MsgBox("The Chart 1 parameter must Temperature.  The Chart 2 parameter must be Oxygen.  The Chart 1 Year and the Chart 2 Year must be the same.", 64)
    ActiveWindow.ScrollColumn = 1
    ActiveWindow.ScrollRow = 1
    Sheets("Lake Probe Data").Range("h3").Select
    Exit Sub
End If
ActiveWindow.ScrollColumn = 1
    ActiveWindow.ScrollRow = 1
ActiveSheet.ChartObjects("Chart 2").Activate
ActiveSheet.Shapes("Chart 2").ZOrder msoBringToFront
ActiveSheet.ChartObjects("Chart 15").Activate
ActiveSheet.Shapes("Chart 15").ZOrder msoBringToFront
End Sub
Private Sub worksheet_Activate()
Worksheets("Lake Probe Data").Activate
    Sheets("Lake Probe Data").Select
    TextBox1.Visible = False
    CommandButton4.Caption = "Open"
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
Private Sub Worksheet_SelectionChange(ByVal Target As Range)    'this code runs if anything is changed
Dim PlotValue(500) As Double
Dim SelectedDate As Date, TestDate As Date
Dim SelectedSite As String
Dim DataCount As Integer
Application.ScreenUpdating = False
If ActiveCell.Address <> "$R$13" And ActiveCell.Address <> "$R$10" And ActiveCell.Address <> "$R$11" And ActiveCell.Address <> "$R$12" And ActiveCell.Address <> "$R$13" And ActiveCell.Address <> "$R$14" And ActiveCell.Address <> "$R$15" And ActiveCell.Address <> "$R$16" And ActiveCell.Address <> "$R$17" And ActiveCell.Address <> "$R$18" And ActiveCell.Address <> "$R$19" And ActiveCell.Address <> "$R$20" And ActiveCell.Address <> "$R$21" And ActiveCell.Address <> "$R$22" And ActiveCell.Address <> "$R$23" And ActiveCell.Address <> "$R$24" And ActiveCell.Address <> "$R$25" And ActiveCell.Address <> "$R$26" And ActiveCell.Address <> "$R$27" And ActiveCell.Address <> "$R$28" And ActiveCell.Address <> "$R$29" And ActiveCell.Address <> "$R$30" Then Exit Sub
If ActiveCell.Value = "" Then Exit Sub
SelectedDate = ActiveCell.Value
Range("R13:R30").Interior.ColorIndex = 2    'make all cells white again after selection
ActiveCell.Interior.ColorIndex = 15
If ActiveCell.Address = "$R$13" Then: Sheets("Lake Probe Data").Range("m43").Select
If ActiveCell.Address = "$R$14" Then: Sheets("Lake Probe Data").Range("m44").Select
If ActiveCell.Address = "$R$15" Then: Sheets("Lake Probe Data").Range("m45").Select
If ActiveCell.Address = "$R$16" Then: Sheets("Lake Probe Data").Range("m46").Select
If ActiveCell.Address = "$R$17" Then: Sheets("Lake Probe Data").Range("m47").Select
If ActiveCell.Address = "$R$18" Then: Sheets("Lake Probe Data").Range("m48").Select
If ActiveCell.Address = "$R$19" Then: Sheets("Lake Probe Data").Range("m49").Select
If ActiveCell.Address = "$R$20" Then: Sheets("Lake Probe Data").Range("m50").Select
If ActiveCell.Address = "$R$21" Then: Sheets("Lake Probe Data").Range("m51").Select
If ActiveCell.Address = "$R$22" Then: Sheets("Lake Probe Data").Range("m52").Select
If ActiveCell.Address = "$R$23" Then: Sheets("Lake Probe Data").Range("m53").Select
If ActiveCell.Address = "$R$24" Then: Sheets("Lake Probe Data").Range("m54").Select
If ActiveCell.Address = "$R$25" Then: Sheets("Lake Probe Data").Range("m55").Select
If ActiveCell.Address = "$R$26" Then: Sheets("Lake Probe Data").Range("m56").Select
If ActiveCell.Address = "$R$27" Then: Sheets("Lake Probe Data").Range("m57").Select
If ActiveCell.Address = "$R$28" Then: Sheets("Lake Probe Data").Range("m58").Select
If ActiveCell.Address = "$R$29" Then: Sheets("Lake Probe Data").Range("m59").Select
If ActiveCell.Address = "$R$30" Then: Sheets("Lake Probe Data").Range("m60").Select
If ActiveCell.Address = "$R$31" Then: Sheets("Lake Probe Data").Range("m61").Select
If ActiveCell.Address = "$R$32" Then: Sheets("Lake Probe Data").Range("m62").Select
If ActiveCell.Address = "$R$33" Then: Sheets("Lake Probe Data").Range("m63").Select
If ActiveCell.Address = "$R$34" Then: Sheets("Lake Probe Data").Range("m63").Select
YJump = Right(ActiveCell.Address, 2)
YDown = 14 + YJump
For i = 1 To 8
        PlotValue(i) = ActiveCell.Value
        ActiveCell.Offset(0, 1).Select
Next i
Sheets("Lake Probe Data").Range("ah43").Select  ''''''''print
For i = 1 To 8
  ActiveCell.Value = PlotValue(i)
  ActiveCell.Offset(0, -1).Select
Next i
ActiveCell.Offset(YDown, -13).Select
For i = 1 To 8
        PlotValue(i) = ActiveCell.Value
        ActiveCell.Offset(0, 1).Select
Next i
Sheets("Lake Probe Data").Range("ah100").Select ''''''''''''''''print
For i = 1 To 8
  ActiveCell.Value = PlotValue(i)
  ActiveCell.Offset(0, -1).Select
Next i
ActiveSheet.ChartObjects("Chart 2").Activate  'y axis
ActiveChart.Axes(xlValue).Select
ActiveChart.Axes(xlValue).MajorUnit = 15
ActiveChart.Axes(xlValue).MaximumScale = 90
ActiveChart.Axes(xlValue).MinimumScale = 0
ActiveSheet.ChartObjects("Chart 2").Activate    'x axis
ActiveChart.Axes(xlCategory).Select
ActiveChart.Axes(xlCategory).MajorUnit = 15
ActiveChart.Axes(xlCategory).MaximumScale = 90
ActiveChart.Axes(xlCategory).MinimumScale = 30
 ActiveSheet.ChartObjects("Chart 15").Activate  'y axis
ActiveChart.Axes(xlValue).Select
ActiveChart.Axes(xlValue).MajorUnit = 15
ActiveChart.Axes(xlValue).MaximumScale = 90
ActiveChart.Axes(xlValue).MinimumScale = 0
ActiveSheet.ChartObjects("Chart 15").Activate    'x axis
ActiveChart.Axes(xlCategory).Select
ActiveChart.Axes(xlCategory).MajorUnit = 2
ActiveChart.Axes(xlCategory).MaximumScale = 16
ActiveChart.Axes(xlCategory).MinimumScale = 0
Sheets("Lake Probe Data").Range("ac13").Select
ActiveCell.Value = "Temperature   " & CStr(SelectedDate)
 ActiveCell.Offset(1, 0).Select
ActiveCell.Value = "Oxygen  " & CStr(SelectedDate)
Sheets("Lake Probe Data").Range("h3").Select
End Sub