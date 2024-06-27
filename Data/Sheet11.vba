Attribute VB_Name = "Sheet11"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Attribute VB_Control = "CommandButton1, 1, 0, MSForms, CommandButton"
Attribute VB_Control = "CommandButton2, 2, 1, MSForms, CommandButton"
Attribute VB_Control = "CommandButton3, 10, 2, MSForms, CommandButton"
Attribute VB_Control = "TextBox1, 639, 3, MSForms, TextBox"
Attribute VB_Control = "CommandButton4, 643, 4, MSForms, CommandButton"
Attribute VB_Control = "CommandButton5, 644, 5, MSForms, CommandButton"
Private Sub CommandButton1_Click()
Dim Data1Count As Integer, Data2Count As Integer
Dim TestDate As Date, TestValue As Double, PMax As Double
Dim PlotDate(4000) As Date, PlotValue(4000) As Double
Dim StartYear As Integer, EndYear As Integer, SelectYear
Dim Site1 As String, Site2 As String, Site3 As String, Site As String
Dim NewSite1 As String, NewSite2 As String, NewSite3 As String
Dim Avg(10) As Variant
'Stream DATA  Chart 1  Left Hand chart   1 year
Application.ScreenUpdating = False
Sheets("Stream Chemistry").Range("i4").Select: SelectYear = ActiveCell.Value    'read selections
Sheets("Stream Chemistry").Range("k3").Select: Site1 = ActiveCell.Value
Sheets("Stream Chemistry").Range("k4").Select: Site2 = ActiveCell.Value
Sheets("Stream Chemistry").Range("k5").Select: Site3 = ActiveCell.Value
Sheets("Stream Chemistry").Range("ad39").Select: NewSite1 = ActiveCell.Value
Sheets("Stream Chemistry").Range("ag39").Select: NewSite2 = ActiveCell.Value
Sheets("Stream Chemistry").Range("aj39").Select: NewSite3 = ActiveCell.Value
'Sheets("Stream Chemistry").Range("aa39").Select: NewSite3 = ActiveCell.Value   'Collision
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If Site1 = "Stone" Or Site2 = "Stone" Or Site3 = "Stone" Then        '1             CHECK TO SEE IF THERE ARE DATA FOR THE SELECTED YEAR
    SelectColumn = 2
    ActiveSheet.Cells(38, SelectColumn).Select: MinYear = ActiveCell.Value  'min
    ActiveSheet.Cells(39, SelectColumn).Select: MaxYear = ActiveCell.Value  'max
    If SelectYear > MaxYear Or SelectYear < MinYear Then
        response1 = MsgBox("Data are not available for Stone for the Year = " & CStr(SelectYear) & ". Want to try different year?", 4)
        If response1 = vbYes Then
            response2 = MsgBox("Enter Year between " & CStr(MinYear) & " and " & CStr(MaxYear) & ".", 0)
            Exit Sub
        End If
    End If
End If
If Site1 = "Vet's" Or Site2 = "Vet's" Or Site3 = "Vet's" Then   '2
    SelectColumn = 5
    ActiveSheet.Cells(38, SelectColumn).Select: MinYear = ActiveCell.Value  'min
    ActiveSheet.Cells(39, SelectColumn).Select: MaxYear = ActiveCell.Value  'max
    If SelectYear > MaxYear Or SelectYear < MinYear Then
        response1 = MsgBox("Data are not available for Vet's for the Year = " & CStr(SelectYear) & ". Want to try different year?", 4)
        If response1 = vbYes Then
        response2 = MsgBox("Enter Year between " & CStr(MinYear) & " and " & CStr(MaxYear) & ".", 0)
        Exit Sub
      End If
    End If
End If
If Site1 = "Haze" Or Site2 = "Haze" Or Site3 = "Haze" Then  '3
    SelectColumn = 8
    ActiveSheet.Cells(38, SelectColumn).Select: MinYear = ActiveCell.Value  'min
    ActiveSheet.Cells(39, SelectColumn).Select: MaxYear = ActiveCell.Value  'max
    If SelectYear > MaxYear Or SelectYear < MinYear Then
        response1 = MsgBox("Data are not available for Haze for the Year = " & CStr(SelectYear) & ". Want to try different year?", 4)
        If response1 = vbYes Then
        response2 = MsgBox("Enter Year between " & CStr(MinYear) & " and " & CStr(MaxYear) & ".", 0)
        Exit Sub
    End If
  End If
End If
If Site1 = "Carter" Or Site2 = "Carter" Or Site3 = "Carter" Then    '4
    SelectColumn = 11
    ActiveSheet.Cells(38, SelectColumn).Select: MinYear = ActiveCell.Value  'min
    ActiveSheet.Cells(39, SelectColumn).Select: MaxYear = ActiveCell.Value  'max
    If SelectYear > MaxYear Or SelectYear < MinYear Then
        response1 = MsgBox("Data are not available for Carter for the Year = " & CStr(SelectYear) & ". Want to try different year?", 4)
        If response1 = vbYes Then
        response2 = MsgBox("Enter Year between " & CStr(MinYear) & " and " & CStr(MaxYear) & ".", 0)
        Exit Sub
    End If
  End If
End If
If Site1 = "Pioneer" Or Site2 = "Pioneer" Or Site3 = "Pioneer" Then '5
    SelectColumn = 14
    ActiveSheet.Cells(38, SelectColumn).Select: MinYear = ActiveCell.Value  'min
    ActiveSheet.Cells(39, SelectColumn).Select: MaxYear = ActiveCell.Value  'max
    If SelectYear > MaxYear Or SelectYear < MinYear Then
        response1 = MsgBox("Data are not available for Pioneer for the Year = " & CStr(SelectYear) & ". Want to try different year?", 4)
        If response1 = vbYes Then
            response2 = MsgBox("Enter Year between " & CStr(MinYear) & " and " & CStr(MaxYear) & ".", 0)
            Exit Sub
        End If
    End If
End If
If Site1 = "USGS" Or Site2 = "USGS" Or Site3 = "USGS" Then  '6
    SelectColumn = 17
    ActiveSheet.Cells(38, SelectColumn).Select: MinYear = ActiveCell.Value  'min
    ActiveSheet.Cells(39, SelectColumn).Select: MaxYear = ActiveCell.Value  'max
    If SelectYear > MaxYear Or SelectYear < MinYear Then
        response1 = MsgBox("Data are not available for USGS for the Year = " & CStr(SelectYear) & ". Want to try different year?", 4)
        If response1 = vbYes Then
        response2 = MsgBox("Enter Year between " & CStr(MinYear) & " and " & CStr(MaxYear) & ".", 0)
        Exit Sub
      End If
    End If
End If
If Site1 = "Ind Hill" Or Site2 = "Ind Hill" Or Site3 = "Ind Hill" Then  '7
    SelectColumn = 20
    ActiveSheet.Cells(38, SelectColumn).Select: MinYear = ActiveCell.Value  'min
    ActiveSheet.Cells(39, SelectColumn).Select: MaxYear = ActiveCell.Value  'max
    If SelectYear > MaxYear Or SelectYear < MinYear Then
        response1 = MsgBox("Data are not available for Ind Hill for the Year = " & CStr(SelectYear) & ". Want to try different year?", 4)
        If response1 = vbYes Then
        response2 = MsgBox("Enter Year between " & CStr(MinYear) & " and " & CStr(MaxYear) & ".", 0)
        Exit Sub
    End If
  End If
End If
If Site1 = "Dead" Or Site2 = "Dead" Or Site3 = "Dead" Then  '8
    SelectColumn = 23
    ActiveSheet.Cells(38, SelectColumn).Select: MinYear = ActiveCell.Value  'min
    ActiveSheet.Cells(39, SelectColumn).Select: MaxYear = ActiveCell.Value  'max
   ' ActiveSheet.Cells(38, SelectColumn+1).Select: Data1Count = ActiveCell.Value  'Data1Count
    If SelectYear > MaxYear Or SelectYear < MinYear Then
        response1 = MsgBox("Data are not available for Dead for the Year = " & CStr(SelectYear) & ". Want to try different year?", 4)
        If response1 = vbYes Then
        response2 = MsgBox("Enter Year between " & CStr(MinYear) & " and " & CStr(MaxYear) & ".", 0)
        Exit Sub
    End If
  End If
End If
'why stop at dead??   add Collision + 3 others
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sheets("Stream Chemistry").Range("an40:as4000").ClearContents 'clear plot data from previous run
If Site1 = "Stone" Then Sheets("Stream Chemistry").Range("c38").Select: Data1Count = ActiveCell.Value
If Site1 = "Vet's" Then Sheets("Stream Chemistry").Range("f38").Select: Data1Count = ActiveCell.Value
If Site1 = "Haze" Then Sheets("Stream Chemistry").Range("i38").Select: Data1Count = ActiveCell.Value
If Site1 = "Carter" Then Sheets("Stream Chemistry").Range("l38").Select: Data1Count = ActiveCell.Value
If Site1 = "Pioneer" Then Sheets("Stream Chemistry").Range("o38").Select: Data1Count = ActiveCell.Value
'collision after poineer
If Site1 = "USGS" Then Sheets("Stream Chemistry").Range("r38").Select: Data1Count = ActiveCell.Value
If Site1 = "Ind Hill" Then Sheets("Stream Chemistry").Range("u38").Select: Data1Count = ActiveCell.Value
If Site1 = "Dead" Then Sheets("Stream Chemistry").Range("x38").Select: Data1Count = ActiveCell.Value
If Site1 = "Collision" Then Sheets("Stream Chemistry").Range("aa38").Select: Data1Count = ActiveCell.Value
If Site1 = NewSite1 Then Sheets("Stream Chemistry").Range("ad38").Select: Data1Count = ActiveCell.Value
If Site1 = NewSite2 Then Sheets("Stream Chemistry").Range("ag38").Select: Data1Count = ActiveCell.Value
If Site1 = NewSite3 Then Sheets("Stream Chemistry").Range("aj38").Select: Data1Count = ActiveCell.Value
ActiveCell.Offset(2, -1).Select     'Move to first date
k = 1
For j = 1 To Data1Count      'begin the scan through all the data
            TestDate = ActiveCell.Value     'select the first date record
            ActiveCell.Offset(0, 1).Select
            TestValue = ActiveCell.Value    'select the first date parameter value
            If Year(TestDate) = SelectYear Then   'And Year(TestDate) <= EndYear Then 'check
                PlotDate(k) = TestDate
                PlotValue(k) = TestValue
            k = k + 1                   'increment the count
            End If
            ActiveCell.Offset(1, -1).Select     'go down one and back to the date column
Next j
Sheets("Stream Chemistry").Range("an40").Select     'print range for plot
For i = 1 To k - 1                                                                      'need to substract one because of k=k+1 above
    dayofyear = DateDiff("d", CDate("1/1/" & Year(PlotDate(i))), PlotDate(i)) + 1
    ActiveCell.Value = dayofyear
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Value = PlotValue(i)
    ActiveCell.Offset(1, -1).Select
Next i
'xxxxxxxxxxxxxxxxxxxxxxx
If Site2 = "Stone" Then Sheets("Stream Chemistry").Range("c38").Select: Data2Count = ActiveCell.Value
If Site2 = "Vet's" Then Sheets("Stream Chemistry").Range("f38").Select: Data2Count = ActiveCell.Value
If Site2 = "Haze" Then Sheets("Stream Chemistry").Range("i38").Select: Data2Count = ActiveCell.Value
If Site2 = "Carter" Then Sheets("Stream Chemistry").Range("l38").Select: Data2Count = ActiveCell.Value
If Site2 = "Pioneer" Then Sheets("Stream Chemistry").Range("o38").Select: Data2Count = ActiveCell.Value
If Site2 = "USGS" Then Sheets("Stream Chemistry").Range("r38").Select: Data2Count = ActiveCell.Value
If Site2 = "Ind Hill" Then Sheets("Stream Chemistry").Range("u38").Select: Data2Count = ActiveCell.Value
If Site2 = "Dead" Then Sheets("Stream Chemistry").Range("x38").Select: Data2Count = ActiveCell.Value
If Site2 = "Collision" Then Sheets("Stream Chemistry").Range("aa38").Select: Data2Count = ActiveCell.Value
If Site2 = NewSite1 Then Sheets("Stream Chemistry").Range("ad38").Select: Data2Count = ActiveCell.Value
If Site2 = NewSite2 Then Sheets("Stream Chemistry").Range("ag38").Select: Data2Count = ActiveCell.Value
If Site2 = NewSite3 Then Sheets("Stream Chemistry").Range("aj38").Select: Data2Count = ActiveCell.Value
ActiveCell.Offset(2, -1).Select
k = 1
For j = 1 To Data2Count      'begin the scan through all the data
            TestDate = ActiveCell.Value     'select the first date record
            ActiveCell.Offset(0, 1).Select
            TestValue = ActiveCell.Value    'select the first date parameter value
            If Year(TestDate) = SelectYear Then 'check
                PlotDate(k) = TestDate
                PlotValue(k) = TestValue
                k = k + 1                   'increment the count
            End If
            ActiveCell.Offset(1, -1).Select     'go down one and back to the date column
Next j
Sheets("Stream Chemistry").Range("ap40").Select     'print range for plot
For i = 1 To k - 1         'need to substract one because of k=k+1 above
    dayofyear = DateDiff("d", CDate("1/1/" & Year(PlotDate(i))), PlotDate(i)) + 1
    ActiveCell.Value = dayofyear
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Value = PlotValue(i)
    ActiveCell.Offset(1, -1).Select
Next i
'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
If Site3 = "Stone" Then Sheets("Stream Chemistry").Range("c38").Select: Data3Count = ActiveCell.Value
If Site3 = "Vet's" Then Sheets("Stream Chemistry").Range("f38").Select: Data3Count = ActiveCell.Value
If Site3 = "Haze" Then Sheets("Stream Chemistry").Range("i38").Select: Data3Count = ActiveCell.Value
If Site3 = "Carter" Then Sheets("Stream Chemistry").Range("l38").Select: Data3Count = ActiveCell.Value
If Site3 = "Pioneer" Then Sheets("Stream Chemistry").Range("o38").Select: Data3Count = ActiveCell.Value
If Site3 = "USGS" Then Sheets("Stream Chemistry").Range("r38").Select: Data3Count = ActiveCell.Value
If Site3 = "Ind Hill" Then Sheets("Stream Chemistry").Range("u38").Select: Data3Count = ActiveCell.Value
If Site3 = "Dead" Then Sheets("Stream Chemistry").Range("x38").Select: Data3Count = ActiveCell.Value
If Site3 = "Collision" Then Sheets("Stream Chemistry").Range("aa38").Select: Data3Count = ActiveCell.Value
If Site3 = NewSite1 Then Sheets("Stream Chemistry").Range("ad38").Select: Data3Count = ActiveCell.Value
If Site3 = NewSite2 Then Sheets("Stream Chemistry").Range("ag38").Select: Data3Count = ActiveCell.Value
If Site3 = NewSite3 Then Sheets("Stream Chemistry").Range("aj38").Select: Data3Count = ActiveCell.Value
ActiveCell.Offset(2, -1).Select
k = 1
For j = 1 To Data3Count      'begin the scan through all the data
            TestDate = ActiveCell.Value     'select the first date record
            ActiveCell.Offset(0, 1).Select
            TestValue = ActiveCell.Value    'select the first date parameter value
            If Year(TestDate) = SelectYear Then 'check
                PlotDate(k) = TestDate
                PlotValue(k) = TestValue
                k = k + 1                   'increment the count
            End If
            ActiveCell.Offset(1, -1).Select     'go down one and back to the date column
Next j
Sheets("Stream Chemistry").Range("ar40").Select     'print range for plot
For i = 1 To k - 1         'need to substract one because of k=k+1 above
    dayofyear = DateDiff("d", CDate("1/1/" & Year(PlotDate(i))), PlotDate(i)) + 1
    ActiveCell.Value = dayofyear
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Value = PlotValue(i)
    ActiveCell.Offset(1, -1).Select
Next i
'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
ActiveSheet.ChartObjects("Chart 8").Activate            'Right click chart and Chart Window
ActiveChart.ChartArea.Select
ActiveChart.Axes(xlCategory).Select
    With ActiveChart.Axes(xlCategory)
        .MinimumScale = 0
        .MaximumScale = 360
    End With
For k = 1 To 3
    If k = 1 Then Sheets("Stream Chemistry").Range("ao37").Select
    If k = 2 Then Sheets("Stream Chemistry").Range("aq37").Select
    If k = 3 Then Sheets("Stream Chemistry").Range("as37").Select
    Avg(k) = ActiveCell.Value
Next k
Sheets("Stream Chemistry").Range("ao18").Select: ActiveCell.Value = " TP Year = " + CStr(SelectYear)    'define chart title
ActiveCell.Offset(1, 0).Select
ActiveCell.Value = Site1 & " = "
ActiveCell.Offset(1, 0).Select
ActiveCell.Value = Site2 & " = "
ActiveCell.Offset(1, 0).Select
ActiveCell.Value = Site3 & " = "
Sheets("Annual Averages").Select        'gives annual average using available data for the year  eeve if year incomplete
                                        ' therefore annual average values change  good!
    For i = 1 To 3
        If i = 1 Then Site = Site1
        If i = 2 Then Site = Site2
        If i = 3 Then Site = Site3
        If Site = "Stone" And SelectYear > 2009 Then
          Sheets("Annual Averages").Range("g" + CStr(SelectYear - 2010 + 48)).Select
          ActiveCell.Value = Avg(i)
        End If
        If Site = "Carter" And SelectYear > 2009 Then
        Sheets("Annual Averages").Range("h" + CStr(SelectYear - 2010 + 48)).Select
        ActiveCell.Value = Avg(i)
        End If
        If Site = "Collision" And SelectYear > 2009 Then
        Sheets("Annual Averages").Range("i" + CStr(SelectYear - 2010 + 48)).Select
        ActiveCell.Value = Avg(i)
        End If
        If Site = "Dead" And SelectYear > 2009 Then
        Sheets("Annual Averages").Range("j" + CStr(SelectYear - 2010 + 48)).Select
        ActiveCell.Value = Avg(i)
        End If
        If Site = "Vet's" And SelectYear > 2009 Then
        Sheets("Annual Averages").Range("k" + CStr(SelectYear - 2010 + 48)).Select
        ActiveCell.Value = Avg(i)
        End If
        If Site = "Pioneer" And SelectYear > 2009 Then
        Sheets("Annual Averages").Range("l" + CStr(SelectYear - 2010 + 48)).Select
        ActiveCell.Value = Avg(i)
        End If
        If Site = "USGS" And SelectYear > 2009 Then
        Sheets("Annual Averages").Range("m" + CStr(SelectYear - 2010 + 48)).Select
        ActiveCell.Value = Avg(i)
        End If
        If Site = "Haze" And SelectYear > 2009 Then
        Sheets("Annual Averages").Range("n" + CStr(SelectYear - 2010 + 48)).Select
        ActiveCell.Value = Avg(i)
        End If
        If Site = "Ind Hill" And SelectYear > 2009 Then                                     'not sure why we need Ind Hill
        Sheets("Annual Averages").Range("o" + CStr(SelectYear - 2010 + 48)).Select
        ActiveCell.Value = Avg(i)
        End If
   Next i
Sheets("Stream Chemistry").Select
Sheets("Stream Chemistry").Range("i4").Select
End Sub
Private Sub CommandButton2_Click()
Dim Data1Count As Integer, Data2Count As Integer
Dim TestDate As Date, TestValue As Double, PMax As Double
Dim PlotDate(4000) As Date, PlotValue(4000) As Double
Dim StartYear As Integer, EndYear As Integer, SelectYear
Dim Site1 As String, Site2 As String, Site3 As String, Site As String
Dim Avg(10) As Variant
'Stream DATA  Chart 2  Left Hand chart   1 year
Application.ScreenUpdating = False
Sheets("Stream Chemistry").Range("o4").Select: StartYear = ActiveCell.Value
Sheets("Stream Chemistry").Range("o5").Select: EndYear = ActiveCell.Value
Sheets("Stream Chemistry").Range("p4").Select: Site1 = ActiveCell.Value
Sheets("Stream Chemistry").Range("p5").Select: Site2 = ActiveCell.Value
Sheets("Stream Chemistry").Range("ad39").Select: NewSite1 = ActiveCell.Value
Sheets("Stream Chemistry").Range("ag39").Select: NewSite2 = ActiveCell.Value
Sheets("Stream Chemistry").Range("aj39").Select: NewSite3 = ActiveCell.Value
If EndYear < StartYear Then
    response% = MsgBox("The End Year must be greater than or equal to the Start Year.", 64)
    Exit Sub
End If
Sheets("Stream Chemistry").Range("at40:aw4000").ClearContents 'clear plot data from previous run
If Site1 = "Stone" Then Sheets("Stream Chemistry").Range("c38").Select: Data1Count = ActiveCell.Value
If Site1 = "Vet's" Then Sheets("Stream Chemistry").Range("f38").Select: Data1Count = ActiveCell.Value
If Site1 = "Haze" Then Sheets("Stream Chemistry").Range("i38").Select: Data1Count = ActiveCell.Value
If Site1 = "Carter" Then Sheets("Stream Chemistry").Range("l38").Select: Data1Count = ActiveCell.Value
If Site1 = "Pioneer" Then Sheets("Stream Chemistry").Range("o38").Select: Data1Count = ActiveCell.Value
If Site1 = "USGS" Then Sheets("Stream Chemistry").Range("r38").Select: Data1Count = ActiveCell.Value
If Site1 = "Ind Hill" Then Sheets("Stream Chemistry").Range("u38").Select: Data1Count = ActiveCell.Value
If Site1 = "Dead" Then Sheets("Stream Chemistry").Range("x38").Select: Data1Count = ActiveCell.Value
If Site1 = "Collision" Then Sheets("Stream Chemistry").Range("aa38").Select: Data1Count = ActiveCell.Value
If Site1 = NewSite1 Then Sheets("Stream Chemistry").Range("ad38").Select: Data1Count = ActiveCell.Value
If Site1 = NewSite2 Then Sheets("Stream Chemistry").Range("ag38").Select: Data1Count = ActiveCell.Value
If Site1 = NewSite3 Then Sheets("Stream Chemistry").Range("aj38").Select: Data1Count = ActiveCell.Value
ActiveCell.Offset(2, -1).Select
k = 1
For j = 1 To Data1Count      'begin the scan through all the data
            TestDate = ActiveCell.Value     'select the first date record
            ActiveCell.Offset(0, 1).Select
            TestValue = ActiveCell.Value    'select the first date parameter value
            If Year(TestDate) >= StartYear And Year(TestDate) <= EndYear Then 'check
                PlotDate(k) = TestDate
                PlotValue(k) = TestValue
            k = k + 1                   'increment the count
            End If
            ActiveCell.Offset(1, -1).Select     'go down one and back to the date column
Next j
Sheets("Stream Chemistry").Range("at40").Select     'print range for plot
For i = 1 To k - 1                                                                      'need to substract one because of k=k+1 above
    dayofyear = DateDiff("d", CDate("1/1/" & Year(PlotDate(i))), PlotDate(i)) + 1
    ActiveCell.Value = Year(PlotDate(i)) + dayofyear / 365
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Value = PlotValue(i)
    ActiveCell.Offset(1, -1).Select
Next i
'xxxxxxxxxxxxxxxxxxxxxxx
If Site2 = "Stone" Then Sheets("Stream Chemistry").Range("c38").Select: Data2Count = ActiveCell.Value
If Site2 = "Vet's" Then Sheets("Stream Chemistry").Range("f38").Select: Data2Count = ActiveCell.Value
If Site2 = "Haze" Then Sheets("Stream Chemistry").Range("i38").Select: Data2Count = ActiveCell.Value
If Site2 = "Carter" Then Sheets("Stream Chemistry").Range("l38").Select: Data2Count = ActiveCell.Value
If Site2 = "Pioneer" Then Sheets("Stream Chemistry").Range("o38").Select: Data2Count = ActiveCell.Value
If Site2 = "USGS" Then Sheets("Stream Chemistry").Range("r38").Select: Data2Count = ActiveCell.Value
If Site2 = "Ind Hill" Then Sheets("Stream Chemistry").Range("u38").Select: Data2Count = ActiveCell.Value
If Site2 = "Dead" Then Sheets("Stream Chemistry").Range("x38").Select: Data2Count = ActiveCell.Value
If Site2 = "Collision" Then Sheets("Stream Chemistry").Range("aa38").Select: Data2Count = ActiveCell.Value
If Site2 = NewSite1 Then Sheets("Stream Chemistry").Range("ad38").Select: Data2Count = ActiveCell.Value
If Site2 = NewSite2 Then Sheets("Stream Chemistry").Range("ag38").Select: Data2Count = ActiveCell.Value
If Site2 = NewSite3 Then Sheets("Stream Chemistry").Range("aj38").Select: Data2Count = ActiveCell.Value
ActiveCell.Offset(2, -1).Select
k = 1
For j = 1 To Data2Count      'begin the scan through all the data
            TestDate = ActiveCell.Value     'select the first date record
            ActiveCell.Offset(0, 1).Select
            TestValue = ActiveCell.Value    'select the first date parameter value
            If Year(TestDate) >= StartYear And Year(TestDate) <= EndYear Then 'check
                PlotDate(k) = TestDate
                PlotValue(k) = TestValue
                k = k + 1                   'increment the count
            End If
            ActiveCell.Offset(1, -1).Select     'go down one and back to the date column
Next j
Sheets("Stream Chemistry").Range("av40").Select     'print range for plot
For i = 1 To k - 1         'need to substract one because of k=k+1 above
    dayofyear = DateDiff("d", CDate("1/1/" & Year(PlotDate(i))), PlotDate(i)) + 1
    ActiveCell.Value = Year(PlotDate(i)) + dayofyear / 365
       ActiveCell.Offset(0, 1).Select
    ActiveCell.Value = PlotValue(i)
    ActiveCell.Offset(1, -1).Select
Next i
YAxisLabel = "mg/m3"
Sheets("Stream Chemistry").Range("ap3").Select: ActiveCell.Value = YAxisLabel
ActiveSheet.ChartObjects("Chart 9").Activate            'Right click chart and Chart Window
ActiveChart.ChartArea.Select
ActiveChart.Axes(xlCategory).Select
    With ActiveChart.Axes(xlCategory)
        .MinimumScale = StartYear
        .MaximumScale = EndYear + 1
    End With
'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
Sheets("Stream Chemistry").Range("at38").Select: ActiveCell.Value = Site1
ActiveCell.Offset(0, 2).Select
ActiveCell.Value = Site2
Sheets("Stream Chemistry").Range("as19").Select: ActiveCell.Value = "Start = " + CStr(StartYear) + "    End = " + CStr(EndYear) ''''''''''''''' chart title
Sheets("Stream Chemistry").Range("o4").Select
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
Private Sub CommandButton5_Click()  'Import Code
Application.ScreenUpdating = False
Dim answer As Integer
Dim NewDate As Date, LastEnteredDate As Date
 '''''''''''''''''''''''''''''''''''  read new values for import
 Sheets("Stream Chemistry").Select
        Sheets("Stream Chemistry").Range("x10").Select
        NewDate = ActiveCell.Value
        ActiveCell.Offset(1, 0).Select
        NewStone = ActiveCell.Value
        ActiveCell.Offset(1, 0).Select
        NewVets = ActiveCell.Value
        ActiveCell.Offset(1, 0).Select
        NewHaze = ActiveCell.Value
        ActiveCell.Offset(1, 0).Select
        NewCarter = ActiveCell.Value
        ActiveCell.Offset(1, 0).Select
        NewPioneer = ActiveCell.Value
        ActiveCell.Offset(1, 0).Select
        NewUSGS = ActiveCell.Value
        ActiveCell.Offset(1, 0).Select
        NewIndHill = ActiveCell.Value
        ActiveCell.Offset(1, 0).Select
        NewDead = ActiveCell.Value
        Sheets("Stream Chemistry").Range("c38").Select  'find last date
            StoneCount = ActiveCell.Value
            Sheets("Stream Chemistry").Range("b" + CStr(StoneCount + 39)).Select
            LastEnteredDate = ActiveCell.Value
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If NewDate = LastEnteredDate Then     'replace old data with new data
            answer = MsgBox("OK to replace the old data on this date?", vbQuestion + vbYesNo)
            If answer = 7 Then Exit Sub 'no
            If answer = 6 Then          'yes
                Sheets("Stream Chemistry").Range("c38").Select  'find stone count & replace
                StoneCount = ActiveCell.Value
                Sheets("Stream Chemistry").Range("c" + CStr(StoneCount + 39)).Select
                ActiveCell.Value = NewStone
                Sheets("Stream Chemistry").Range("f38").Select  'find Vets count & replace
                VetsCount = ActiveCell.Value
                Sheets("Stream Chemistry").Range("f" + CStr(VetsCount + 39)).Select
                ActiveCell.Value = NewVets
                Sheets("Stream Chemistry").Range("i38").Select  'find Haze count & replace
                HazeCount = ActiveCell.Value
                Sheets("Stream Chemistry").Range("i" + CStr(HazeCount + 39)).Select
                ActiveCell.Value = NewHaze
                Sheets("Stream Chemistry").Range("L38").Select  'find Carter count & replace
                CarterCount = ActiveCell.Value
                Sheets("Stream Chemistry").Range("L" + CStr(CarterCount + 39)).Select
                ActiveCell.Value = NewCarter
                Sheets("Stream Chemistry").Range("o38").Select  'find Pioneer count & replace
                PioneerCount = ActiveCell.Value
                Sheets("Stream Chemistry").Range("o" + CStr(PioneerCount + 39)).Select
                ActiveCell.Value = NewPioneer
                Sheets("Stream Chemistry").Range("r38").Select  'find USGS count & replace
                USGSCount = ActiveCell.Value
                Sheets("Stream Chemistry").Range("r" + CStr(USGSCount + 39)).Select
                ActiveCell.Value = NewUSGS
                Sheets("Stream Chemistry").Range("u38").Select  'find IndHill count & replace
                IndHillCount = ActiveCell.Value
                Sheets("Stream Chemistry").Range("u" + CStr(IndHillCount + 39)).Select
                ActiveCell.Value = NewIndHill
                Sheets("Stream Chemistry").Range("x38").Select  'find Dead count & replace
                DeadCount = ActiveCell.Value
                Sheets("Stream Chemistry").Range("x" + CStr(DeadCount + 39)).Select
                ActiveCell.Value = NewDead
                answer = MsgBox("Replacement Completed.", 64)
            End If
        End If
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If NewDate > LastEnteredDate Then     'new import
            answer = MsgBox("OK to import new data for this date?", vbQuestion + vbYesNo)
            If answer = 7 Then Exit Sub
            If answer = 6 Then          'yes
                Sheets("Stream Chemistry").Range("c38").Select  'find stone count & replace
                StoneCount = ActiveCell.Value
                Sheets("Stream Chemistry").Range("c" + CStr(StoneCount + 40)).Select
                ActiveCell.Value = NewStone
                ActiveCell.Offset(0, -1).Select
                ActiveCell.Value = NewDate
                Sheets("Stream Chemistry").Range("f38").Select  'find Vets count & replace
                VetsCount = ActiveCell.Value
                Sheets("Stream Chemistry").Range("f" + CStr(VetsCount + 40)).Select
                ActiveCell.Value = NewVets
                ActiveCell.Offset(0, -1).Select
                ActiveCell.Value = NewDate
                Sheets("Stream Chemistry").Range("i38").Select  'find Haze count & replace
                HazeCount = ActiveCell.Value
                Sheets("Stream Chemistry").Range("i" + CStr(HazeCount + 40)).Select
                ActiveCell.Value = NewHaze
                ActiveCell.Offset(0, -1).Select
                ActiveCell.Value = NewDate
                Sheets("Stream Chemistry").Range("L38").Select  'find Carter count & replace
                CarterCount = ActiveCell.Value
                Sheets("Stream Chemistry").Range("L" + CStr(CarterCount + 40)).Select
                ActiveCell.Value = NewCarter
                ActiveCell.Offset(0, -1).Select
                ActiveCell.Value = NewDate
                Sheets("Stream Chemistry").Range("o38").Select  'find Pioneer count & replace
                PioneerCount = ActiveCell.Value
                Sheets("Stream Chemistry").Range("o" + CStr(PioneerCount + 40)).Select
                ActiveCell.Value = NewPioneer
                ActiveCell.Offset(0, -1).Select
                ActiveCell.Value = NewDate
                Sheets("Stream Chemistry").Range("r38").Select  'find USGS count & replace
                USGSCount = ActiveCell.Value
                Sheets("Stream Chemistry").Range("r" + CStr(USGSCount + 40)).Select
                ActiveCell.Value = NewUSGS
                ActiveCell.Offset(0, -1).Select
                ActiveCell.Value = NewDate
                Sheets("Stream Chemistry").Range("u38").Select  'find IndHill count & replace
                IndHillCount = ActiveCell.Value
                Sheets("Stream Chemistry").Range("u" + CStr(IndHillCount + 40)).Select
                ActiveCell.Value = NewIndHill
                ActiveCell.Offset(0, -1).Select
                ActiveCell.Value = NewDate
                Sheets("Stream Chemistry").Range("x38").Select  'find Dead count & replace
                DeadCount = ActiveCell.Value
                Sheets("Stream Chemistry").Range("x" + CStr(DeadCount + 40)).Select
                ActiveCell.Value = NewDead
                ActiveCell.Offset(0, -1).Select
                ActiveCell.Value = NewDate
                answer = MsgBox("New Data Entered.", 64)
            End If
        End If
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If NewDate < LastEnteredDate Then     'deny
            answer = MsgBox("Cannot replace the old data on this date.", 64)
            Exit Sub
        End If
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Sub
Private Sub worksheet_Activate()
Worksheets("Stream Chemistry").Activate
    Sheets("Stream Chemistry").Select
    TextBox1.Visible = False
    CommandButton4.Caption = "Open"
End Sub