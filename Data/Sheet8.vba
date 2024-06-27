Attribute VB_Name = "Sheet8"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Attribute VB_Control = "CommandButton4, 754, 0, MSForms, CommandButton"
Attribute VB_Control = "CommandButton1, 1, 1, MSForms, CommandButton"
Attribute VB_Control = "CommandButton2, 2, 2, MSForms, CommandButton"
Attribute VB_Control = "CommandButton3, 10, 3, MSForms, CommandButton"
Attribute VB_Control = "TextBox1, 755, 4, MSForms, TextBox"
Attribute VB_Control = "CommandButton5, 756, 5, MSForms, CommandButton"
Private Sub CommandButton1_Click()
Dim MeasureDate(10000) As Date, TestDate As Date, PlotDate1(100) As Date, PlotDate2(100) As Date, PlotDate3(100) As Date, PlotDate4(100) As Date, PlotDate5(100) As Date
Dim TPDate(10000) As Date, TPSur(10000) As Double, TPMid(10000) As Double, TPBot(10000) As Double, TPVolWt(10000) As Double
Dim DataCount As Integer, Parameter As String, SelectedYear As Integer
Dim PlotValue1(100) As Double, PlotValue2(100), PlotValue3(100), PlotValue4(100), PlotValue5(100)
Dim AnnualAverage As Double
Dim PMax As Double
Dim TestDepth As String
Dim Halt As Boolean
Dim NewTestDate(5000) As Date, NewTestDepth(5000) As String, NewTestValue(5000) As Double, VolumeWeightedValue(5000) As Double
Dim ForwardValue As Double
'LAKE CHEMISTRY  Chart 1
Application.ScreenUpdating = False
Sheets("Lake Chemistry").Select
Sheets("Lake Chemistry").Range("aq4:aq90").ClearContents 'clear plot data from previous run         'clear previous results
Sheets("Lake Chemistry").Range("as4:as90").ClearContents
Sheets("Lake Chemistry").Range("at4:at90").ClearContents
Sheets("Lake Chemistry").Range("av4:av90").ClearContents
Sheets("Lake Chemistry").Range("aw4:aw90").ClearContents
Sheets("Lake Chemistry").Range("ay4:ay90").ClearContents
Sheets("Lake Chemistry").Range("bk4:bk90").ClearContents
Sheets("Lake Chemistry").Range("bm4:bm90").ClearContents
Sheets("Lake Chemistry").Range("h3").Select: SelectedYear = ActiveCell.Value    'year
Sheets("Lake Chemistry").Range("h4").Select: Parameter = ActiveCell.Value       'Parm
If Parameter = "Total P" Then SelectColumn = 2
If Parameter = "Nitrate" Then SelectColumn = 7
If Parameter = "Chlorophyll" Then SelectColumn = 10
If Parameter = "Secchi" Then SelectColumn = 13
If Parameter = "TDP" Then SelectColumn = 16
ActiveSheet.Cells(37, SelectColumn).Select: MinYear = ActiveCell.Value  'min
ActiveSheet.Cells(38, SelectColumn).Select: MaxYear = ActiveCell.Value  'max
If SelectedYear > MaxYear Or SelectedYear < MinYear Then
    response% = MsgBox("Data are not available for Year = " & CStr(SelectedYear) & " as entered in Cell H3.", 64)
    response% = MsgBox("Enter Year between " & CStr(MinYear) & " and " & CStr(MaxYear) & ".", 64)
    Exit Sub
End If
If Parameter <> "Total P" Then
If Parameter = "Nitrate" Then Sheets("Lake Chemistry").Range("i37").Select: DataCount = ActiveCell.Value
If Parameter = "Chlorophyll" Then Sheets("Lake Chemistry").Range("L37").Select: DataCount = ActiveCell.Value
If Parameter = "Secchi" Then Sheets("Lake Chemistry").Range("o37").Select: DataCount = ActiveCell.Value
If Parameter = "TDP" Then Sheets("Lake Chemistry").Range("r37").Select: DataCount = ActiveCell.Value
        k = 0   'count for number of dates during the year where data are available for surface
        kk = 0   'count for number of dates during the year where data are available for middle
        kkk = 0     'count for number of dates during the year where data are available for bottom
        If Parameter = "Nitrate" Then Sheets("Lake Chemistry").Range("g39").Select        'first date
        If Parameter = "Chlorophyll" Then Sheets("Lake Chemistry").Range("j39").Select    'first date
        If Parameter = "Secchi" Then Sheets("Lake Chemistry").Range("m39").Select         'first date
        If Parameter = "TDP" Then Sheets("Lake Chemistry").Range("p39").Select            'first date
        For j = 1 To DataCount      'begin the scan through all the data
            TestDate = ActiveCell.Value     'select the first date record
            ActiveCell.Offset(0, 1).Select
            TestDepth = ActiveCell.Value    'select the first depth record
            ActiveCell.Offset(0, 1).Select
            TestValue = ActiveCell.Value    'select the first date parameter value
            'xxxxxxxxxxxxxxxxxx
            If Year(TestDate) = SelectedYear And TestDepth = "Sur" Then 'check to verify that record is for the selected year and depth
                PlotDate1(k) = TestDate
                PlotValue1(k) = TestValue
                k = k + 1                   'increment the count
            End If
            'xxxxxxxxxxxxxxxxxxxxxxx
            If Year(TestDate) = SelectedYear And TestDepth = "Mid" Then 'check to verify that record is for the selected year and depth
                PlotDate2(kk) = TestDate
                PlotValue2(kk) = TestValue
                kk = kk + 1                   'increment the count
            End If
            'xxxxxxxxxxxxxxxxxxxxxxx
            If Year(TestDate) = SelectedYear And TestDepth = "Bot" Then    'check to verify that record is for the selected year and depth
                PlotDate3(kkk) = TestDate
                PlotValue3(kkk) = TestValue
                kkk = kkk + 1                   'increment the count
            End If
            ActiveCell.Offset(1, -2).Select     'go down one and back to the date column
        Next j
        Sheets("Lake Chemistry").Range("aq4").Select     'print range for plot for 3 depths
        For i = 1 To k
            ActiveCell.Value = PlotDate1(i)
            ActiveCell.Offset(0, 2).Select
            ActiveCell.Value = PlotValue1(i)
            ActiveCell.Offset(1, -2).Select
        Next i
        Sheets("Lake Chemistry").Range("at4").Select     'print range for plot
        For i = 1 To kk
            ActiveCell.Value = PlotDate2(i)
            ActiveCell.Offset(0, 2).Select
            ActiveCell.Value = PlotValue2(i)
            ActiveCell.Offset(1, -2).Select
        Next i
        Sheets("Lake Chemistry").Range("aw4").Select     'print range for plot
        For i = 1 To kkk
            ActiveCell.Value = PlotDate3(i)
            ActiveCell.Offset(0, 2).Select
            ActiveCell.Value = PlotValue3(i)
            ActiveCell.Offset(1, -2).Select
        Next i
If Parameter <> "Secchi" Then YAxisLabel = "mg/m3"         'Y axis label for 4 different parameters
If Parameter = "Secchi" Then YAxisLabel = "feet"
Sheets("Lake Chemistry").Range("ai5").Select
ActiveCell.Value = YAxisLabel
ActiveSheet.ChartObjects("Chart 8").Activate            'Right click chart and Chart Window
ActiveChart.ChartArea.Select
If Parameter = "Nitrate" Then
        With ActiveChart.Axes(xlValue)
            .MinimumScale = 0
            .MaximumScale = 500
            .MajorUnit = 100
        End With
End If
If Parameter = "Chlorophyll" Then                       'vertical scale for 4 different Parms
        With ActiveChart.Axes(xlValue)
            .MinimumScale = 0
            .MaximumScale = 6
            .MajorUnit = 1
        End With
End If
If Parameter = "Secchi" Then
        With ActiveChart.Axes(xlValue)
            .MinimumScale = 0
            .MaximumScale = 40
            .MajorUnit = 10
        End With
End If
If Parameter = "TDP" Then
        With ActiveChart.Axes(xlValue)
            .MinimumScale = 0
            .MaximumScale = 16
            .MajorUnit = 4
        End With
End If
If Parameter = "Secchi" Then
    Sheets("Lake Chemistry").Range("ah12").Select: AnnualAverage = Round(ActiveCell.Value, 1)
    Sheets("Lake Chemistry").Range("ai4").Select: ActiveCell.Value = Parameter + " " + CStr(SelectedYear) + " Average = " + CStr(AnnualAverage)    'define chart title
End If
If Parameter <> "Secchi" Then Sheets("Lake Chemistry").Range("ai4").Select: ActiveCell.Value = Parameter + " " + CStr(SelectedYear)     'define chart title
End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If Parameter = "Total P" Then
    Sheets("Lake Chemistry").Range("bk3:bk90").ClearContents
    Sheets("Lake Chemistry").Range("bm3:bm90").ClearContents
    Sheets("Lake Chemistry").Range("f37").Select: DataCount = ActiveCell.Value
    Sheets("Lake Chemistry").Range("b39").Select        'first date
    For i = 1 To DataCount
        TPDate(i) = ActiveCell.Value     'select the first date record
        ActiveCell.Offset(0, 1).Select      'read all data
        TPSur(i) = ActiveCell.Value
        ActiveCell.Offset(0, 1).Select
        TPMid(i) = ActiveCell.Value
        ActiveCell.Offset(0, 1).Select
        TPBot(i) = ActiveCell.Value
        ActiveCell.Offset(0, 1).Select
        TPVolWt(i) = ActiveCell.Value
        ActiveCell.Offset(1, -4).Select
    Next i
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Sheets("Lake Chemistry").Range("aq4").Select     'print range for plot for 3 depths
    For i = 1 To DataCount
        If Year(TPDate(i)) = SelectedYear Then          'print just values for selected year
            ActiveCell.Value = TPDate(i)
            ActiveCell.Offset(0, 2).Select
            ActiveCell.Value = TPSur(i)
            ActiveCell.Offset(1, -2).Select
        End If
    Next i
    Sheets("Lake Chemistry").Range("at4").Select     'print range for plot for 3 depths
    For i = 1 To DataCount
        If Year(TPDate(i)) = SelectedYear Then
            ActiveCell.Value = TPDate(i)
            ActiveCell.Offset(0, 2).Select
            ActiveCell.Value = TPMid(i)
            ActiveCell.Offset(1, -2).Select
        End If
    Next i
    Sheets("Lake Chemistry").Range("aw4").Select     'print range for plot for 3 depths
    For i = 1 To DataCount
        If Year(TPDate(i)) = SelectedYear Then
            ActiveCell.Value = TPDate(i)
            ActiveCell.Offset(0, 2).Select
            ActiveCell.Value = TPBot(i)
            ActiveCell.Offset(1, -2).Select
        End If
    Next i
     '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' all the data for AQ to AY entered
    Sheets("Lake Chemistry").Range("bk4").Select     'print range for plot
    For i = 1 To DataCount
         If Year(TPDate(i)) = SelectedYear And Year(TPDate(i - 1)) < SelectedYear Then          'find date & value before
            ValueBeforeOne = TPVolWt(i - 1)
            DateBeforeOne = TPDate(i - 1)
            DayBeforeOne = DateDiff("d", CDate("1/1/" & Year(DateBeforeOne)), DateBeforeOne) + 1
            FirstValue = TPVolWt(i)
            FirstDay = DateDiff("d", CDate("1/1/" & Year(TPDate(i))), TPDate(i)) + 1
         End If
           If Year(TPDate(i)) = SelectedYear And Year(TPDate(i + 1)) > SelectedYear Then              'find date & value after
            ValueAfter365 = TPVolWt(i + 1)
            DateAfter365 = TPDate(i + 1)
            DayAfter365 = DateDiff("d", CDate("1/1/" & Year(DateAfter365)), DateAfter365) + 1
            Lastvalue = TPVolWt(i)
            LastDay = DateDiff("d", CDate("1/1/" & Year(TPDate(i))), TPDate(i)) + 1
         End If
         If Year(TPDate(i)) = SelectedYear Then                                                     'print measured values
            ActiveCell.Value = TPDate(i)
            ActiveCell.Offset(0, 2).Select
            ActiveCell.Value = TPVolWt(i)
            ActiveCell.Offset(1, -2).Select
         End If
    Next i
    DayOneValue = ValueBeforeOne + ((FirstValue - ValueBeforeOne) / (FirstDay - DayBeforeOne + 365)) * (365 - DayBeforeOne + 1)   'checked ok 6-7-2022
    day365value = Lastvalue + (ValueAfter365 - Lastvalue) / (DayAfter365 + 365 - LastDay) * (365 - LastDay)  'checked ok   6-7-2022
   If day365value = 0 Then day365value = TPVolWt(i - 1)
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Sheets("Lake Chemistry").Range("bn2").Select: PointCount = ActiveCell.Value    'this section inserts day 1 and day 365 date & values in correct location
    Sheets("Lake Chemistry").Range("bk3").Select: ActiveCell.Value = DateSerial(SelectedYear, 1, 1)
    ActiveCell.Offset(0, 2).Select
    ActiveCell.Value = DayOneValue
    Sheets("Lake Chemistry").Range("bk" + CStr(PointCount + 4)).Select
    ActiveCell.Value = DateSerial(SelectedYear, 12, 31)
    ActiveCell.Offset(0, 2).Select
    ActiveCell.Value = day365value
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Sheets("Lake Chemistry").Range("ap1").Select            'chart scale  & vertical title
    PMax = ActiveCell.Value
    ActiveSheet.ChartObjects("Chart 8").Activate            'Right click chart and Chart Window
    ActiveChart.ChartArea.Select
    If Parameter = "Total P" And PMax > 16 Then
            With ActiveChart.Axes(xlValue)
                .MinimumScale = 0
                .MaximumScale = 24
                .MajorUnit = 4
            End With
    End If
    If Parameter = "Total P" And PMax <= 16 Then
            With ActiveChart.Axes(xlValue)
                .MinimumScale = 0
                .MaximumScale = 16
                .MajorUnit = 4
            End With
    End If
    Sheets("Lake Chemistry").Range("ai5").Select: ActiveCell.Value = "mg/m3"
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Sheets("Lake Chemistry").Range("bn2").Select     'count number of fixed values except for last
    PointCount = ActiveCell.Value
    Call Macro2           'find days > 8
    Sheets("Lake Chemistry").Range("bx3").Select        'just the measurements
    AnnualAverage = Round(ActiveCell.Value, 2)
    Sheets("Lake Chemistry").Range("ah14").Select
    AnnualAverageExtended = Round(ActiveCell.Value, 2)  'includes measurements plus both tails
    'get both here
    Sheets("Lake Chemistry").Range("bn3").Select: Days = ActiveCell.Value
    Sheets("Lake Chemistry").Range("ai4").Select: ActiveCell.Value = Parameter + " " + CStr(SelectedYear) + " Average = " + CStr(AnnualAverageExtended) + "  " + "Compliance % = " + CStr(Round(100 * (365 - Days) / 365, 0))
    If SelectedYear >= 2010 And Parameter = "Total P" Then       'wait until the end of the year to send to annual averages
    'If SelectedYear >= 2010 And Parameter = "Total P" And SelectedYear < Year(Now()) Then       'wait until the end of the year to send to annual averages
        Sheets("Annual Averages").Select
        Sheets("Annual Averages").Range("c" + CStr(SelectedYear - 2010 + 48)).Select
        ActiveCell.Value = AnnualAverage                                                        'just the measurements
        ActiveCell.Offset(0, 1).Select
        ActiveCell.Value = 100 * (365 - Days) / 365
        Sheets("Annual Averages").Range("d" + CStr(SelectedYear - 2010 + 152)).Select
        ActiveCell.Value = AnnualAverage                                                        'just the measurements
        ActiveCell.Offset(0, 1).Select
        ActiveCell.Value = 100 * (365 - Days) / 365
        Sheets("Lake Chemistry").Select
    End If
End If
Sheets("Lake Chemistry").Range("h4").Select
End Sub
Private Sub CommandButton2_Click()
Dim MeasureDate(10000) As Date, TestDate As Date, PlotDate1(100) As Date, PlotDate2(100) As Date, PlotDate3(100) As Date, PlotDate4(100) As Date, PlotDate5(100) As Date
Dim TPDate(10000) As Date, TPSur(10000) As Double, TPMid(10000) As Double, TPBot(10000) As Double, TPVolWt(10000) As Double
Dim DataCount As Integer, Parameter As String, SelectedYear As Integer
Dim PlotValue1(100) As Double, PlotValue2(100), PlotValue3(100), PlotValue4(100), PlotValue5(100)
Dim AnnualAverage As Double
Dim PMax As Double
Dim TestDepth As String
Dim Halt As Boolean
Dim NewTestDate(5000) As Date, NewTestDepth(5000) As String, NewTestValue(5000) As Double, VolumeWeightedValue(5000) As Double
Dim ForwardValue As Double
'LAKE CHEMISTRY  Chart 2
Application.ScreenUpdating = False
Sheets("Lake Chemistry").Select
Sheets("Lake Chemistry").Range("ba4:ba90").ClearContents 'clear plot data from previous run         'clear previous results
Sheets("Lake Chemistry").Range("bc4:bc90").ClearContents
Sheets("Lake Chemistry").Range("bd4:bd90").ClearContents
Sheets("Lake Chemistry").Range("bf4:bf90").ClearContents
Sheets("Lake Chemistry").Range("bg4:bg90").ClearContents
Sheets("Lake Chemistry").Range("bi4:bi90").ClearContents
Sheets("Lake Chemistry").Range("bq4:bq90").ClearContents
Sheets("Lake Chemistry").Range("bs4:bs90").ClearContents
Sheets("Lake Chemistry").Range("n3").Select: SelectedYear = ActiveCell.Value    'year
Sheets("Lake Chemistry").Range("n4").Select: Parameter = ActiveCell.Value       'Parm
If Parameter = "Total P" Then SelectColumn = 2
If Parameter = "Nitrate" Then SelectColumn = 7
If Parameter = "Chlorophyll" Then SelectColumn = 10
If Parameter = "Secchi" Then SelectColumn = 13
If Parameter = "TDP" Then SelectColumn = 16
ActiveSheet.Cells(37, SelectColumn).Select: MinYear = ActiveCell.Value  'min
ActiveSheet.Cells(38, SelectColumn).Select: MaxYear = ActiveCell.Value  'max
If SelectedYear > MaxYear Or SelectedYear < MinYear Then
    response% = MsgBox("Data are not available for Year = " & CStr(SelectedYear) & " as entered in Cell N3.", 64)
    response% = MsgBox("Enter Year between " & CStr(MinYear) & " and " & CStr(MaxYear) & ".", 64)
    Exit Sub
End If
If Parameter <> "Total P" Then
If Parameter = "Nitrate" Then Sheets("Lake Chemistry").Range("i37").Select: DataCount = ActiveCell.Value
If Parameter = "Chlorophyll" Then Sheets("Lake Chemistry").Range("l37").Select: DataCount = ActiveCell.Value
If Parameter = "Secchi" Then Sheets("Lake Chemistry").Range("o37").Select: DataCount = ActiveCell.Value
If Parameter = "TDP" Then Sheets("Lake Chemistry").Range("r37").Select: DataCount = ActiveCell.Value
        k = 0   'count for number of dates during the year where data are available for surface
        kk = 0   'count for number of dates during the year where data are available for middle
        kkk = 0     'count for number of dates during the year where data are available for bottom
        If Parameter = "Nitrate" Then Sheets("Lake Chemistry").Range("g39").Select        'first date
        If Parameter = "Chlorophyll" Then Sheets("Lake Chemistry").Range("j39").Select        'first date
        If Parameter = "Secchi" Then Sheets("Lake Chemistry").Range("m39").Select        'first date
        If Parameter = "TDP" Then Sheets("Lake Chemistry").Range("p39").Select        'first date
        For j = 1 To DataCount      'begin the scan through all the data
            TestDate = ActiveCell.Value     'select the first date record
            ActiveCell.Offset(0, 1).Select
            TestDepth = ActiveCell.Value    'select the first depth record
            ActiveCell.Offset(0, 1).Select
            TestValue = ActiveCell.Value    'select the first date parameter value
            'xxxxxxxxxxxxxxxxxx
            If Year(TestDate) = SelectedYear And TestDepth = "Sur" Then 'check to verify that record is for the selected year and depth
                PlotDate1(k) = TestDate
                PlotValue1(k) = TestValue
                k = k + 1                   'increment the count
            End If
            'xxxxxxxxxxxxxxxxxxxxxxx
            If Year(TestDate) = SelectedYear And TestDepth = "Mid" Then 'check to verify that record is for the selected year and depth
                PlotDate2(kk) = TestDate
                PlotValue2(kk) = TestValue
                kk = kk + 1                   'increment the count
            End If
            'xxxxxxxxxxxxxxxxxxxxxxx
            If Year(TestDate) = SelectedYear And TestDepth = "Bot" Then    'check to verify that record is for the selected year and depth
                PlotDate3(kkk) = TestDate
                PlotValue3(kkk) = TestValue
                kkk = kkk + 1                   'increment the count
            End If
            ActiveCell.Offset(1, -2).Select     'go down one and back to the date column
        Next j
        Sheets("Lake Chemistry").Range("ba4").Select     'print range for plot for 3 depths
        For i = 1 To k
            ActiveCell.Value = PlotDate1(i)
            ActiveCell.Offset(0, 2).Select
            ActiveCell.Value = PlotValue1(i)
            ActiveCell.Offset(1, -2).Select
        Next i
        Sheets("Lake Chemistry").Range("bd4").Select     'print range for plot
        For i = 1 To kk
            ActiveCell.Value = PlotDate2(i)
            ActiveCell.Offset(0, 2).Select
            ActiveCell.Value = PlotValue2(i)
            ActiveCell.Offset(1, -2).Select
        Next i
        Sheets("Lake Chemistry").Range("bg4").Select     'print range for plot
        For i = 1 To kkk
            ActiveCell.Value = PlotDate3(i)
            ActiveCell.Offset(0, 2).Select
            ActiveCell.Value = PlotValue3(i)
            ActiveCell.Offset(1, -2).Select
        Next i
If Parameter <> "Secchi" Then YAxisLabel = "mg/m3"         'Y axis label for 4 different parameters
If Parameter = "Secchi" Then YAxisLabel = "feet"
Sheets("Lake Chemistry").Range("ai10").Select
ActiveCell.Value = YAxisLabel
ActiveSheet.ChartObjects("Chart 3").Activate            'Right click chart and Chart Window
ActiveChart.ChartArea.Select
If Parameter = "Nitrate" Then
        With ActiveChart.Axes(xlValue)
            .MinimumScale = 0
            .MaximumScale = 500
            .MajorUnit = 100
        End With
End If
If Parameter = "Chlorophyll" Then                       'vertical scale for 4 different Parms
        With ActiveChart.Axes(xlValue)
            .MinimumScale = 0
            .MaximumScale = 6
            .MajorUnit = 1
        End With
End If
If Parameter = "Secchi" Then
        With ActiveChart.Axes(xlValue)
            .MinimumScale = 0
            .MaximumScale = 40
            .MajorUnit = 10
        End With
End If
If Parameter = "TDP" Then
        With ActiveChart.Axes(xlValue)
            .MinimumScale = 0
            .MaximumScale = 16
            .MajorUnit = 4
        End With
End If
If Parameter = "Secchi" Then
    Sheets("Lake Chemistry").Range("ah13").Select: AnnualAverage = Round(ActiveCell.Value, 1)
    Sheets("Lake Chemistry").Range("ai9").Select: ActiveCell.Value = Parameter + " " + CStr(SelectedYear) + " Average = " + CStr(AnnualAverage)    'define chart title
End If
If Parameter <> "Secchi" Then Sheets("Lake Chemistry").Range("ai9").Select: ActiveCell.Value = Parameter + " " + CStr(SelectedYear)     'define chart title
End If
'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
If Parameter = "Total P" Then
    Sheets("Lake Chemistry").Range("bq3:bq90").ClearContents
    Sheets("Lake Chemistry").Range("bs3:bs90").ClearContents
    Sheets("Lake Chemistry").Range("f37").Select: DataCount = ActiveCell.Value
    Sheets("Lake Chemistry").Range("b39").Select        'first date
    For i = 1 To DataCount
        TPDate(i) = ActiveCell.Value     'select the first date record
        ActiveCell.Offset(0, 1).Select
        TPSur(i) = ActiveCell.Value
        ActiveCell.Offset(0, 1).Select
        TPMid(i) = ActiveCell.Value
        ActiveCell.Offset(0, 1).Select
        TPBot(i) = ActiveCell.Value
        ActiveCell.Offset(0, 1).Select
        TPVolWt(i) = ActiveCell.Value
        ActiveCell.Offset(1, -4).Select
    Next i
    Sheets("Lake Chemistry").Range("ba4").Select     'print range for plot for 3 depths
    For i = 1 To DataCount
        If Year(TPDate(i)) = SelectedYear Then
            ActiveCell.Value = TPDate(i)
            ActiveCell.Offset(0, 2).Select
            ActiveCell.Value = TPSur(i)
            ActiveCell.Offset(1, -2).Select
        End If
    Next i
    Sheets("Lake Chemistry").Range("bd4").Select     'print range for plot for 3 depths
    For i = 1 To DataCount
        If Year(TPDate(i)) = SelectedYear Then
            ActiveCell.Value = TPDate(i)
            ActiveCell.Offset(0, 2).Select
            ActiveCell.Value = TPMid(i)
            ActiveCell.Offset(1, -2).Select
        End If
    Next i
    Sheets("Lake Chemistry").Range("bg4").Select     'print range for plot for 3 depths
    For i = 1 To DataCount
        If Year(TPDate(i)) = SelectedYear Then
            ActiveCell.Value = TPDate(i)
            ActiveCell.Offset(0, 2).Select
            ActiveCell.Value = TPBot(i)
            ActiveCell.Offset(1, -2).Select
        End If
    Next i
    Sheets("Lake Chemistry").Range("bq4").Select     'print range for plot
    For i = 1 To DataCount
         If Year(TPDate(i)) = SelectedYear And Year(TPDate(i - 1)) < SelectedYear Then
            ValueBeforeOne = TPVolWt(i - 1)
            DateBeforeOne = TPDate(i - 1)
            DayBeforeOne = DateDiff("d", CDate("1/1/" & Year(DateBeforeOne)), DateBeforeOne) + 1
            FirstValue = TPVolWt(i)
            FirstDay = DateDiff("d", CDate("1/1/" & Year(TPDate(i))), TPDate(i)) + 1
         End If
         If Year(TPDate(i)) = SelectedYear And Year(TPDate(i + 1)) > SelectedYear Then
            ValueAfter365 = TPVolWt(i + 1)
            DateAfter365 = TPDate(i + 1)
            DayAfter365 = DateDiff("d", CDate("1/1/" & Year(DateAfter365)), DateAfter365) + 1
            Lastvalue = TPVolWt(i)
            LastDay = DateDiff("d", CDate("1/1/" & Year(TPDate(i))), TPDate(i)) + 1
         End If
         If Year(TPDate(i)) = SelectedYear Then
            ActiveCell.Value = TPDate(i)
            ActiveCell.Offset(0, 2).Select
            ActiveCell.Value = TPVolWt(i)
            ActiveCell.Offset(1, -2).Select
         End If
    Next i
    DayOneValue = ValueBeforeOne + ((FirstValue - ValueBeforeOne) / (FirstDay - DayBeforeOne + 365)) * (365 - DayBeforeOne + 1)
    day365value = Lastvalue + ((ValueAfter365 - Lastvalue) / (DayAfter365 - LastDay + 365)) * (365 - LastDay)
     If day365value = 0 Then day365value = TPVolWt(i - 1)
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Sheets("Lake Chemistry").Range("bt2").Select: PointCount = ActiveCell.Value
    Sheets("Lake Chemistry").Range("bq3").Select: ActiveCell.Value = DateSerial(SelectedYear, 1, 1)
    ActiveCell.Offset(0, 2).Select
    ActiveCell.Value = DayOneValue
    Sheets("Lake Chemistry").Range("bq" + CStr(PointCount + 4)).Select
    ActiveCell.Value = DateSerial(SelectedYear, 12, 31)
    ActiveCell.Offset(0, 2).Select
    ActiveCell.Value = day365value
    Sheets("Lake Chemistry").Range("az1").Select
    PMax = ActiveCell.Value
    ActiveSheet.ChartObjects("Chart 3").Activate          'Right click chart and Chart Window
    ActiveChart.ChartArea.Select
    If Parameter = "Total P" And PMax > 16 Then
            With ActiveChart.Axes(xlValue)
                .MinimumScale = 0
                .MaximumScale = 24
                .MajorUnit = 4
            End With
    End If
    If Parameter = "Total P" And PMax <= 16 Then
            With ActiveChart.Axes(xlValue)
                .MinimumScale = 0
                .MaximumScale = 16
                .MajorUnit = 4
            End With
    End If
    Sheets("Lake Chemistry").Range("ai10").Select: ActiveCell.Value = "mg/m3"
    Call Macro4           'find days > 8
    Sheets("Lake Chemistry").Range("by3").Select
    AnnualAverage = Round(ActiveCell.Value, 2)
    Sheets("Lake Chemistry").Range("ah15").Select
    AnnualAverageExtended = Round(ActiveCell.Value, 2)
    'get both here
    Sheets("Lake Chemistry").Range("bt3").Select: Days = ActiveCell.Value
    Sheets("Lake Chemistry").Range("ai9").Select: ActiveCell.Value = Parameter + " " + CStr(SelectedYear) + " Average = " + CStr(AnnualAverageExtended) + "  " + "Compliance % = " + CStr(Round(100 * (365 - Days) / 365, 0)) 'define chart title
    If SelectedYear >= 2010 And Parameter = "Total P" And SelectedYear < Year(Now()) Then       'wait until the end of the year to send to annual averages
        Sheets("Annual Averages").Select
        Sheets("Annual Averages").Range("c" + CStr(SelectedYear - 2010 + 48)).Select
        ActiveCell.Value = AnnualAverage                                                        'just the measurements
        ActiveCell.Offset(0, 1).Select
        ActiveCell.Value = 100 * (365 - Days) / 365
        Sheets("Annual Averages").Range("d" + CStr(SelectedYear - 2010 + 152)).Select
        ActiveCell.Value = AnnualAverage                                                        'just the measurements
        ActiveCell.Offset(0, 1).Select
        ActiveCell.Value = 100 * (365 - Days) / 365
        Sheets("Lake Chemistry").Select
    End If
End If
Sheets("Lake Chemistry").Range("n4").Select
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
Private Sub CommandButton5_Click()      'IMPORT DATA
Application.ScreenUpdating = False
Dim answer As Integer
Dim NewDate As Date, LastEnteredTPDate As Date
 '''''''''''''''''''''''''''''''''''  read new values for import
 Sheets("Lake Chemistry").Select
        Sheets("Lake Chemistry").Range("v9").Select
        NewDate = ActiveCell.Value
        ActiveCell.Offset(1, 0).Select
        NewSur = ActiveCell.Value
        ActiveCell.Offset(1, 0).Select
        NewMid = ActiveCell.Value
        ActiveCell.Offset(1, 0).Select
        NewBot = ActiveCell.Value
        ActiveCell.Offset(1, 0).Select
        NewSecchi = ActiveCell.Value
        Sheets("Lake Chemistry").Range("f37").Select: TPCount = ActiveCell.Value        'get TP count
        Sheets("Lake Chemistry").Range("b" + CStr(TPCount + 38)).Select
        LastEnteredTPDate = ActiveCell.Value                                            'get last date
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If NewDate = LastEnteredTPDate Then     'replace?
            answer = MsgBox("OK to replace the old data on this date?", vbQuestion + vbYesNo)
            If answer = 7 Then Exit Sub 'no
            If answer = 6 Then          'yes
                Sheets("Lake Chemistry").Range("c" + CStr(TPCount + 38)).Select
                ActiveCell.Value = NewSur
                ActiveCell.Offset(0, 1).Select
                ActiveCell.Value = NewMid
                ActiveCell.Offset(0, 1).Select
                ActiveCell.Value = NewBot
                Sheets("Lake Chemistry").Range("o37").Select: SecchiCount = ActiveCell.Value    'get secchi count
                Sheets("Lake Chemistry").Range("o" + CStr(SecchiCount + 38)).Select
                ActiveCell.Value = NewSecchi
                answer = MsgBox("Replacement Completed.", 64)
            End If
        End If
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If NewDate > LastEnteredTPDate Then     'new import
            answer = MsgBox("OK to import new data for this date?", vbQuestion + vbYesNo)
            If answer = 7 Then Exit Sub 'no
            If answer = 6 Then          'yes
                Sheets("Lake Chemistry").Range("b" + CStr(TPCount + 39)).Select
                ActiveCell.Value = NewDate
                ActiveCell.Offset(0, 1).Select
                ActiveCell.Value = NewSur
                ActiveCell.Offset(0, 1).Select
                ActiveCell.Value = NewMid
                ActiveCell.Offset(0, 1).Select
                ActiveCell.Value = NewBot
                Sheets("Lake Chemistry").Range("o37").Select: SecchiCount = ActiveCell.Value
                Sheets("Lake Chemistry").Range("m" + CStr(SecchiCount + 39)).Select
                ActiveCell.Value = NewDate
                ActiveCell.Offset(0, 1).Select
                ActiveCell.Value = "Sur"
                ActiveCell.Offset(0, 1).Select
                ActiveCell.Value = NewSecchi
                answer = MsgBox("New Data Entered.", 64)
            End If
        End If
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If NewDate < LastEnteredTPDate Then     'deny
            answer = MsgBox("Cannot replace the old data on this date.", 64)
            Exit Sub
        End If
End Sub
Private Sub worksheet_Activate()
Worksheets("Lake Chemistry").Activate
    Sheets("Lake Chemistry").Select
    TextBox1.Visible = False
    CommandButton4.Caption = "Open"
End Sub