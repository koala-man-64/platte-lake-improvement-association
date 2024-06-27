Attribute VB_Name = "Sheet7"

Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"

Attribute VB_GlobalNameSpace = False

Attribute VB_Creatable = False

Attribute VB_PredeclaredId = True

Attribute VB_Exposed = True

Attribute VB_TemplateDerived = False

Attribute VB_Customizable = True

Attribute VB_Control = "TextBox1, 1885, 0, MSForms, TextBox"

Attribute VB_Control = "CommandButton4, 1884, 1, MSForms, CommandButton"

Attribute VB_Control = "CommandButton1, 1032, 2, MSForms, CommandButton"

Attribute VB_Control = "CommandButton2, 1033, 3, MSForms, CommandButton"

Attribute VB_Control = "CommandButton3, 1127, 4, MSForms, CommandButton"

Private Sub CommandButton1_Click()



Dim TestDate(10000) As Date, TestValue(10000) As Variant, TestDepth(10000) As String

Dim PlotDate(10000) As Date, PlotValue(10000) As Double

Dim StartYear As Integer, EndYear As Integer

Dim DataCount As Integer, Parameter As String

Dim FlowAvg(10000) As Variant, FlowMax(10000) As Variant, FlowMin(10000) As Variant, FlowSD(10000) As Variant



'LONG TERM TREND CODE  Chart 1



Application.ScreenUpdating = False

Sheets("Long-Term Trends").Range("c47:ek138").ClearContents 'clear plot data from previous run xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx



Sheets("Long-Term Trends").Select

Sheets("Long-Term Trends").Range("h3").Select: StartYear = ActiveCell.Value

Sheets("Long-Term Trends").Range("h4").Select: EndYear = ActiveCell.Value

Sheets("Long-Term Trends").Range("h5").Select: Parameter = ActiveCell.Value



If EndYear < StartYear Then

    response% = MsgBox("The End Year must be greater than or equal to the Start Year.", 64)

    Sheets("Long-Term Trends").Range("h3").Select

    Exit Sub

End If



If Parameter = "USGS Flow" Then

        ActiveSheet.ChartObjects("Chart 1030").Visible = False

        ActiveSheet.ChartObjects("Chart 3").Visible = True

        

        Sheets("Long-Term Trends").Range("c267:ci271").ClearContents

        

        Sheets("Flow & Rain Data").Select

        

        FlowCol = 25 + (StartYear - 1990)

        ActiveSheet.Cells(751, FlowCol).Select

        For j = StartYear To EndYear

            FlowAvg(j) = ActiveCell.Value

            

            If FlowAvg(j) = 0 Or FlowAvg(j) = "" Then

                response% = MsgBox("Flow data for one of the selected years have not been entered.", 64)

                Sheets("Long-Term Trends").Select

                Sheets("Long-Term Trends").Range("h4").Select

                Exit Sub

            End If

            

            ActiveCell.Offset(0, 1).Select

        Next j



        FlowCol = 25 + (StartYear - 1990)

        ActiveSheet.Cells(752, FlowCol).Select

        For j = StartYear To EndYear

            FlowMax(j) = ActiveCell.Value

            ActiveCell.Offset(0, 1).Select

        Next j



        FlowCol = 25 + (StartYear - 1990)

        ActiveSheet.Cells(753, FlowCol).Select

        For j = StartYear To EndYear

            FlowMin(j) = ActiveCell.Value

            ActiveCell.Offset(0, 1).Select

        Next j



        FlowCol = 25 + (StartYear - 1990)

        ActiveSheet.Cells(754, FlowCol).Select

        For j = StartYear To EndYear

            FlowSD(j) = ActiveCell.Value

            ActiveCell.Offset(0, 1).Select

        Next j



        Sheets("Long-Term Trends").Select   '''''''''''''''''''''''''''''''''''''''''

        Sheets("Long-Term Trends").Range("c267").Select

        For j = StartYear To EndYear

            ActiveCell.Value = j

            ActiveCell.Offset(0, 1).Select

        Next j



        Sheets("Long-Term Trends").Range("c268").Select

        For j = StartYear To EndYear

            ActiveCell.Value = FlowAvg(j)

            ActiveCell.Offset(0, 1).Select

        Next j



        Sheets("Long-Term Trends").Range("c269").Select

        For j = StartYear To EndYear

            ActiveCell.Value = FlowMax(j)

            ActiveCell.Offset(0, 1).Select

        Next j



        Sheets("Long-Term Trends").Range("c270").Select

        For j = StartYear To EndYear

            ActiveCell.Value = FlowMin(j)

            ActiveCell.Offset(0, 1).Select

        Next j

        

        Sheets("Long-Term Trends").Range("c271").Select

        For j = StartYear To EndYear

            ActiveCell.Value = FlowSD(j)

            ActiveCell.Offset(0, 1).Select

        Next j

        

        Exit Sub

     Else

        ActiveSheet.ChartObjects("Chart 1030").Visible = True

        ActiveSheet.ChartObjects("Chart 3").Visible = False

End If  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



ActiveSheet.ChartObjects("Chart 1030").Activate

ActiveChart.Axes(xlCategory).Select

    

    If EndYear - StartYear <= 5 Then ActiveChart.Axes(xlCategory).MajorUnit = 1

    If EndYear - StartYear > 5 And EndYear - StartYear <= 10 Then ActiveChart.Axes(xlCategory).MajorUnit = 2

    If EndYear - StartYear > 10 Then ActiveChart.Axes(xlCategory).MajorUnit = 5

    

    ActiveChart.ChartArea.Select

    

    ActiveChart.Axes(xlCategory).Select

    With ActiveChart.Axes(xlCategory)

        If StartYear >= 1992 Then .MinimumScale = StartYear - 1

        If StartYear < 1992 Then .MinimumScale = 1990

        .MaximumScale = EndYear + 1

    End With



If Parameter = "Vol Wt Summer Temp" Then

        YAxisLabel = "degrees F"

        With ActiveChart.Axes(xlValue)

            .MinimumScale = 20

            .MaximumScale = 100

            .MajorUnit = 10

        End With

End If

   

If Parameter = "Lake Vol Wt TP" Or Parameter = "Stone TP" Or Parameter = "Vet's TP" Or Parameter = "Haze TP" Or Parameter = "Carter TP" Or Parameter = "Pioneer TP" Or Parameter = "USGS TP" Or Parameter = "NB Ind Hill TP" Or Parameter = "NB Dead TP" Then

        YAxisLabel = "mg/m3"

        With ActiveChart.Axes(xlValue)

            .MinimumScale = 0

            .MaximumScale = 25

            .MajorUnit = 5

        End With

End If



If Parameter = "Chlorophyll" Then

        YAxisLabel = "mg/m3"

        With ActiveChart.Axes(xlValue)

            .MinimumScale = 0

            .MaximumScale = 10

            .MajorUnit = 2

        End With

End If





If Parameter = "Secchi" Then

        YAxisLabel = "feet"

        With ActiveChart.Axes(xlValue)

            .MinimumScale = 0

            .MaximumScale = 35

            .MajorUnit = 5

        End With

End If

  

Sheets("Long-Term Trends").Range("ae5").Select: ActiveCell.Value = YAxisLabel



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

If Parameter = "Lake Vol Wt TP" Then

    Sheets("Lake Chemistry").Select

    Sheets("Lake Chemistry").Range("f37").Select: DataCount = ActiveCell.Value

    Sheets("Lake Chemistry").Range("b39").Select

    For i = 1 To DataCount

        TestDate(i) = ActiveCell.Value

        ActiveCell.Offset(0, 4).Select

        TestValue(i) = ActiveCell.Value

        ActiveCell.Offset(1, -4).Select

    Next i

End If ''''''''''''''''''''''''''''''''''''''''

    

If Parameter = "Chlorophyll" Then

    Sheets("Lake Chemistry").Select

    Sheets("Lake Chemistry").Range("L37").Select: DataCount = ActiveCell.Value

    Sheets("Lake Chemistry").Range("k39").Select

    

    j = 1

    

    For i = 1 To DataCount

    

        TestDepth(i) = ActiveCell.Value

        

        If TestDepth(i) = "Sur" Then

        

            ActiveCell.Offset(0, -1).Select

            TestDate(j) = ActiveCell.Value

            ActiveCell.Offset(0, 2).Select

            TestValue(j) = ActiveCell.Value

            

            j = j + 1

            

            ActiveCell.Offset(0, -1).Select

        

        End If

    

        ActiveCell.Offset(1, 0).Select

    

    Next i

End If ''''''''''''''''''''''''''''''''''''''''

    

If Parameter = "Secchi" Then

    Sheets("Lake Chemistry").Select

    Sheets("Lake Chemistry").Range("o37").Select: DataCount = ActiveCell.Value

    Sheets("Lake Chemistry").Range("m39").Select

    For i = 1 To DataCount

        TestDate(i) = ActiveCell.Value

        ActiveCell.Offset(0, 2).Select

        TestValue(i) = ActiveCell.Value

        ActiveCell.Offset(1, -2).Select

     Next i

End If '''''''''''''''''''''''''''''''''''''''''''''xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx



If Parameter = "Stone TP" Then

    Sheets("Stream Chemistry").Select

    Sheets("Stream Chemistry").Range("c38").Select: DataCount = ActiveCell.Value

    Sheets("Stream Chemistry").Range("b40").Select

    For i = 1 To DataCount

        TestDate(i) = ActiveCell.Value

        ActiveCell.Offset(0, 1).Select

        TestValue(i) = ActiveCell.Value

        ActiveCell.Offset(1, -1).Select

    Next i '''''''''''''''''''''''''''''''''''''''''''''''''

End If



If Parameter = "Vet's TP" Then

    Sheets("Stream Chemistry").Select

    Sheets("Stream Chemistry").Range("f38").Select: DataCount = ActiveCell.Value

    Sheets("Stream Chemistry").Range("e40").Select

    For i = 1 To DataCount

        TestDate(i) = ActiveCell.Value

        ActiveCell.Offset(0, 1).Select

        TestValue(i) = ActiveCell.Value

        ActiveCell.Offset(1, -1).Select

    Next i '''''''''''''''''''''''''''''''''''''''''''''''''

End If



If Parameter = "Haze TP" Then

    Sheets("Stream Chemistry").Select

    Sheets("Stream Chemistry").Range("i38").Select: DataCount = ActiveCell.Value

    Sheets("Stream Chemistry").Range("h40").Select

    For i = 1 To DataCount

        TestDate(i) = ActiveCell.Value

        ActiveCell.Offset(0, 1).Select

        TestValue(i) = ActiveCell.Value

        ActiveCell.Offset(1, -1).Select

    Next i '''''''''''''''''''''''''''''''''''''''''''''''''

End If



If Parameter = "Carter TP" Then

    Sheets("Stream Chemistry").Select

    Sheets("Stream Chemistry").Range("l38").Select: DataCount = ActiveCell.Value

    Sheets("Stream Chemistry").Range("k40").Select

    For i = 1 To DataCount

        TestDate(i) = ActiveCell.Value

        ActiveCell.Offset(0, 1).Select

        TestValue(i) = ActiveCell.Value

        ActiveCell.Offset(1, -1).Select

    Next i '''''''''''''''''''''''''''''''''''''''''''''''''

End If



If Parameter = "Pioneer TP" Then

    Sheets("Stream Chemistry").Select

    Sheets("Stream Chemistry").Range("o38").Select: DataCount = ActiveCell.Value

    Sheets("Stream Chemistry").Range("n40").Select

    For i = 1 To DataCount

        TestDate(i) = ActiveCell.Value

        ActiveCell.Offset(0, 1).Select

        TestValue(i) = ActiveCell.Value

        ActiveCell.Offset(1, -1).Select

    Next i '''''''''''''''''''''''''''''''''''''''''''''''''

End If



If Parameter = "USGS TP" Then

    Sheets("Stream Chemistry").Select

    Sheets("Stream Chemistry").Range("r38").Select: DataCount = ActiveCell.Value

    Sheets("Stream Chemistry").Range("q40").Select

    For i = 1 To DataCount

        TestDate(i) = ActiveCell.Value

        ActiveCell.Offset(0, 1).Select

        TestValue(i) = ActiveCell.Value

        ActiveCell.Offset(1, -1).Select

    Next i '''''''''''''''''''''''''''''''''''''''''''''''''''

End If



If Parameter = "NB Ind Hill TP" Then

    Sheets("Stream Chemistry").Select

    Sheets("Stream Chemistry").Range("u38").Select: DataCount = ActiveCell.Value

    Sheets("Stream Chemistry").Range("t40").Select

    For i = 1 To DataCount

        TestDate(i) = ActiveCell.Value

        ActiveCell.Offset(0, 1).Select

        TestValue(i) = ActiveCell.Value

        ActiveCell.Offset(1, -1).Select

    Next i '''''''''''''''''''''''''''''''''''''''''''''''''

End If



If Parameter = "NB Dead TP" Then

    Sheets("Stream Chemistry").Select

    Sheets("Stream Chemistry").Range("x38").Select: DataCount = ActiveCell.Value

    Sheets("Stream Chemistry").Range("w40").Select

     

   For i = 1 To DataCount

        TestDate(i) = ActiveCell.Value

        ActiveCell.Offset(0, 1).Select

        TestValue(i) = ActiveCell.Value

        ActiveCell.Offset(1, -1).Select

    Next i

End If '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx



If Parameter = "Vol Wt Summer Temp" Then

    Sheets("Lake Probe Data").Select

    Sheets("Lake Probe Data").Range("y42").Select: DataCount = ActiveCell.Value

    Sheets("Lake Probe Data").Range("x43").Select

  

         For i = 1 To DataCount

       TestDate(i) = ActiveCell.Value

       ActiveCell.Offset(0, 1).Select

        TestValue(i) = ActiveCell.Value

        ActiveCell.Offset(1, -1).Select

        

    Next i

End If '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''







Sheets("Long-Term Trends").Select

Sheets("Long-Term Trends").Range("c47").Select

For j = StartYear To EndYear

    ActiveCell.Value = j

    ActiveCell.Offset(0, 1).Select

Next j



Sheets("Long-Term Trends").Range("c48").Select



For k = StartYear To EndYear



    For i = 1 To DataCount

        If Year(TestDate(i)) = StartYear Then

            ActiveCell.Value = TestValue(i)

            ActiveCell.Offset(1, 0).Select

        End If

    Next i

       

    StartYear = StartYear + 1

    

    RowNumber = ActiveCell.Row

    ColumnNumber = ActiveCell.Column

    

    ActiveSheet.Cells(48, ActiveCell.Column + 1).Select



Next k

   

Sheets("Long-Term Trends").Range("h3").Select

    

 '''''''''''''''''''''''''''''''''''''

 

Sheets("Long-Term Trends").Range("ae13").Select 'find calculated value of average (many digets)

LongTermAverage = Round(ActiveCell.Value, 2)    'round

If ActiveCell.Value < 1 Then LongTermAverage = Round(ActiveCell.Value, 3)   'round

Sheets("Long-Term Trends").Range("ae4").Select: ActiveCell.Value = "Average " + Parameter + " = " + CStr(LongTermAverage) + "  " 'create chart title



Sheets("Long-Term Trends").Range("h3").Select

    

End Sub

Private Sub CommandButton2_Click()



'LONG TERM TREND CODE  Chart 2



Dim TestDate(10000) As Date, TestValue(10000) As Variant, TestDepth(10000) As Variant

Dim PlotDate(10000) As Date, PlotValue(10000) As Double

Dim StartYear As Integer, EndYear As Integer

Dim DataCount As Integer, Parameter As String

Dim FlowAvg(10000) As Variant, FlowMax(10000) As Variant, FlowMin(10000) As Variant, FlowSD(10000) As Variant



Application.ScreenUpdating = False

Sheets("Long-Term Trends").Range("c154:ek254").ClearContents 'clear plot data from previous run xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx



Sheets("Long-Term Trends").Select

Sheets("Long-Term Trends").Range("n3").Select: StartYear = ActiveCell.Value

Sheets("Long-Term Trends").Range("n4").Select: EndYear = ActiveCell.Value

Sheets("Long-Term Trends").Range("n5").Select: Parameter = ActiveCell.Value



If EndYear < StartYear Then

    response% = MsgBox("The End Year must be greater than or equal to the Start Year.", 64)

    Sheets("Long-Term Trends").Range("n3").Select

    Exit Sub

End If



If Parameter = "USGS Flow" Then

        ActiveSheet.ChartObjects("Chart 1130").Visible = False

        ActiveSheet.ChartObjects("Chart 12").Visible = True

        

        Sheets("Long-Term Trends").Range("c282:ci286").ClearContents

        

        Sheets("Flow & Rain Data").Select

        

        FlowCol = 25 + (StartYear - 1990)

        ActiveSheet.Cells(751, FlowCol).Select

        For j = StartYear To EndYear

            FlowAvg(j) = ActiveCell.Value

            

            If FlowAvg(j) = 0 Or FlowAvg(j) = "" Then

                response% = MsgBox("Flow data for one of the selected years have not been entered.", 64)

                Sheets("Long-Term Trends").Select

                Sheets("Long-Term Trends").Range("h4").Select

                Exit Sub

            End If

            

            ActiveCell.Offset(0, 1).Select

        Next j



        FlowCol = 25 + (StartYear - 1990)

        ActiveSheet.Cells(752, FlowCol).Select

        For j = StartYear To EndYear

            FlowMax(j) = ActiveCell.Value

            ActiveCell.Offset(0, 1).Select

        Next j



        FlowCol = 25 + (StartYear - 1990)

        ActiveSheet.Cells(753, FlowCol).Select

        For j = StartYear To EndYear

            FlowMin(j) = ActiveCell.Value

            ActiveCell.Offset(0, 1).Select

        Next j



        FlowCol = 25 + (StartYear - 1990)

        ActiveSheet.Cells(754, FlowCol).Select

        For j = StartYear To EndYear

            FlowSD(j) = ActiveCell.Value

            ActiveCell.Offset(0, 1).Select

        Next j



        Sheets("Long-Term Trends").Select   '''''''''''''''''''''''''''''''''''''''''

        Sheets("Long-Term Trends").Range("c282").Select

        For j = StartYear To EndYear

            ActiveCell.Value = j

            ActiveCell.Offset(0, 1).Select

        Next j



        Sheets("Long-Term Trends").Range("c283").Select

        For j = StartYear To EndYear

            ActiveCell.Value = FlowAvg(j)

            ActiveCell.Offset(0, 1).Select

        Next j



        Sheets("Long-Term Trends").Range("c284").Select

        For j = StartYear To EndYear

            ActiveCell.Value = FlowMax(j)

            ActiveCell.Offset(0, 1).Select

        Next j



        Sheets("Long-Term Trends").Range("c285").Select

        For j = StartYear To EndYear

            ActiveCell.Value = FlowMin(j)

            ActiveCell.Offset(0, 1).Select

        Next j

        

        Sheets("Long-Term Trends").Range("c286").Select

        For j = StartYear To EndYear

            ActiveCell.Value = FlowSD(j)

            ActiveCell.Offset(0, 1).Select

        Next j

        

        Exit Sub

     Else

        ActiveSheet.ChartObjects("Chart 1130").Visible = True

        ActiveSheet.ChartObjects("Chart 12").Visible = False

End If  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

ActiveSheet.ChartObjects("Chart 1130").Activate

ActiveChart.Axes(xlCategory).Select

    

    If EndYear - StartYear <= 5 Then ActiveChart.Axes(xlCategory).MajorUnit = 1

    If EndYear - StartYear > 5 And EndYear - StartYear <= 10 Then ActiveChart.Axes(xlCategory).MajorUnit = 2

    If EndYear - StartYear > 10 Then ActiveChart.Axes(xlCategory).MajorUnit = 5

    

    ActiveChart.ChartArea.Select

    

    ActiveChart.Axes(xlCategory).Select

    With ActiveChart.Axes(xlCategory)

        If StartYear >= 1992 Then .MinimumScale = StartYear - 1

        If StartYear < 1992 Then .MinimumScale = 1990

        .MaximumScale = EndYear + 1

    End With



If Parameter = "Vol Wt Summer Temp" Then

        YAxisLabel = "degrees F"

        With ActiveChart.Axes(xlValue)

            .MinimumScale = 20

            .MaximumScale = 100

            .MajorUnit = 10

        End With

End If

   

If Parameter = "Lake Vol Wt TP" Or Parameter = "Stone TP" Or Parameter = "Vet's TP" Or Parameter = "Haze TP" Or Parameter = "Carter TP" Or Parameter = "Pioneer TP" Or Parameter = "USGS TP" Or Parameter = "NB Ind Hill TP" Or Parameter = "NB Dead TP" Then

        YAxisLabel = "mg/m3"

        With ActiveChart.Axes(xlValue)

            .MinimumScale = 0

            .MaximumScale = 25

            .MajorUnit = 5

        End With

End If



If Parameter = "Chlorophyll" Then

        YAxisLabel = "mg/m3"

        With ActiveChart.Axes(xlValue)

            .MinimumScale = 0

            .MaximumScale = 10

            .MajorUnit = 2

        End With

End If



If Parameter = "Secchi" Then

        YAxisLabel = "feet"

        With ActiveChart.Axes(xlValue)

            .MinimumScale = 0

            .MaximumScale = 35

            .MajorUnit = 5

        End With

End If

  

Sheets("Long-Term Trends").Range("ae9").Select: ActiveCell.Value = YAxisLabel

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

If Parameter = "Lake Vol Wt TP" Then

    Sheets("Lake Chemistry").Select

    Sheets("Lake Chemistry").Range("f37").Select: DataCount = ActiveCell.Value

    Sheets("Lake Chemistry").Range("b39").Select

    For i = 1 To DataCount

        TestDate(i) = ActiveCell.Value

        ActiveCell.Offset(0, 4).Select

        TestValue(i) = ActiveCell.Value

        ActiveCell.Offset(1, -4).Select

    Next i

End If ''''''''''''''''''''''''''''''''''''''''

    

 If Parameter = "Chlorophyll" Then

    Sheets("Lake Chemistry").Select

    Sheets("Lake Chemistry").Range("L37").Select: DataCount = ActiveCell.Value

    Sheets("Lake Chemistry").Range("k39").Select

    

    j = 1

    

    For i = 1 To DataCount

    

        TestDepth(i) = ActiveCell.Value

        

        If TestDepth(i) = "Sur" Then

        

            ActiveCell.Offset(0, -1).Select

            TestDate(j) = ActiveCell.Value

            ActiveCell.Offset(0, 2).Select

            TestValue(j) = ActiveCell.Value

            

            j = j + 1

            

            ActiveCell.Offset(0, -1).Select

        

        End If

    

        ActiveCell.Offset(1, 0).Select

    

    Next i

End If '''''''''''''''''''''''''''''''''''''''

        

If Parameter = "Secchi" Then

    Sheets("Lake Chemistry").Select

    Sheets("Lake Chemistry").Range("o37").Select: DataCount = ActiveCell.Value

    Sheets("Lake Chemistry").Range("m39").Select

    For i = 1 To DataCount

        TestDate(i) = ActiveCell.Value

        ActiveCell.Offset(0, 2).Select

        TestValue(i) = ActiveCell.Value

        ActiveCell.Offset(1, -2).Select

     Next i

End If '''''''''''''''''''''''''''''''''''''''''''''



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx



If Parameter = "Stone TP" Then

    Sheets("Stream Chemistry").Select

    Sheets("Stream Chemistry").Range("c38").Select: DataCount = ActiveCell.Value

    Sheets("Stream Chemistry").Range("b40").Select

    For i = 1 To DataCount

        TestDate(i) = ActiveCell.Value

        ActiveCell.Offset(0, 1).Select

        TestValue(i) = ActiveCell.Value

        ActiveCell.Offset(1, -1).Select

    Next i '''''''''''''''''''''''''''''''''''''''''''''''''

End If



If Parameter = "Vet's TP" Then

    Sheets("Stream Chemistry").Select

    Sheets("Stream Chemistry").Range("f38").Select: DataCount = ActiveCell.Value

    Sheets("Stream Chemistry").Range("e40").Select

    For i = 1 To DataCount

        TestDate(i) = ActiveCell.Value

        ActiveCell.Offset(0, 1).Select

        TestValue(i) = ActiveCell.Value

        ActiveCell.Offset(1, -1).Select

    Next i '''''''''''''''''''''''''''''''''''''''''''''''''

End If



If Parameter = "Haze TP" Then

    Sheets("Stream Chemistry").Select

    Sheets("Stream Chemistry").Range("i38").Select: DataCount = ActiveCell.Value

    Sheets("Stream Chemistry").Range("h40").Select

    For i = 1 To DataCount

        TestDate(i) = ActiveCell.Value

        ActiveCell.Offset(0, 1).Select

        TestValue(i) = ActiveCell.Value

        ActiveCell.Offset(1, -1).Select

    Next i '''''''''''''''''''''''''''''''''''''''''''''''''

End If



If Parameter = "Carter TP" Then

    Sheets("Stream Chemistry").Select

    Sheets("Stream Chemistry").Range("l38").Select: DataCount = ActiveCell.Value

    Sheets("Stream Chemistry").Range("k40").Select

    For i = 1 To DataCount

        TestDate(i) = ActiveCell.Value

        ActiveCell.Offset(0, 1).Select

        TestValue(i) = ActiveCell.Value

        ActiveCell.Offset(1, -1).Select

    Next i '''''''''''''''''''''''''''''''''''''''''''''''''

End If



If Parameter = "Pioneer TP" Then

    Sheets("Stream Chemistry").Select

    Sheets("Stream Chemistry").Range("o38").Select: DataCount = ActiveCell.Value

    Sheets("Stream Chemistry").Range("n40").Select

    For i = 1 To DataCount

        TestDate(i) = ActiveCell.Value

        ActiveCell.Offset(0, 1).Select

        TestValue(i) = ActiveCell.Value

        ActiveCell.Offset(1, -1).Select

    Next i '''''''''''''''''''''''''''''''''''''''''''''''''

End If



If Parameter = "USGS TP" Then

    Sheets("Stream Chemistry").Select

    Sheets("Stream Chemistry").Range("r38").Select: DataCount = ActiveCell.Value

    Sheets("Stream Chemistry").Range("q40").Select

    For i = 1 To DataCount

        TestDate(i) = ActiveCell.Value

        ActiveCell.Offset(0, 1).Select

        TestValue(i) = ActiveCell.Value

        ActiveCell.Offset(1, -1).Select

    Next i '''''''''''''''''''''''''''''''''''''''''''''''''''

End If



If Parameter = "NB Ind Hill TP" Then

    Sheets("Stream Chemistry").Select

    Sheets("Stream Chemistry").Range("u38").Select: DataCount = ActiveCell.Value

    Sheets("Stream Chemistry").Range("t40").Select

    For i = 1 To DataCount

        TestDate(i) = ActiveCell.Value

        ActiveCell.Offset(0, 1).Select

        TestValue(i) = ActiveCell.Value

        ActiveCell.Offset(1, -1).Select

    Next i '''''''''''''''''''''''''''''''''''''''''''''''''

End If



If Parameter = "NB Dead TP" Then

    Sheets("Stream Chemistry").Select

    Sheets("Stream Chemistry").Range("x38").Select: DataCount = ActiveCell.Value

    Sheets("Stream Chemistry").Range("w40").Select

     

   For i = 1 To DataCount

        TestDate(i) = ActiveCell.Value

        ActiveCell.Offset(0, 1).Select

        TestValue(i) = ActiveCell.Value

        ActiveCell.Offset(1, -1).Select

    Next i

End If '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx





If Parameter = "Vol Wt Summer Temp" Then

    Sheets("Lake Probe Data").Select

    Sheets("Lake Probe Data").Range("y42").Select: DataCount = ActiveCell.Value

    Sheets("Lake Probe Data").Range("x43").Select

         For i = 1 To DataCount

       TestDate(i) = ActiveCell.Value

       ActiveCell.Offset(0, 1).Select

        TestValue(i) = ActiveCell.Value

        ActiveCell.Offset(1, -1).Select

        

    Next i

End If '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



Sheets("Long-Term Trends").Select

Sheets("Long-Term Trends").Range("c154").Select

For j = StartYear To EndYear

    ActiveCell.Value = j

    ActiveCell.Offset(0, 1).Select

Next j



Sheets("Long-Term Trends").Range("c155").Select



For k = StartYear To EndYear



    For i = 1 To DataCount

        If Year(TestDate(i)) = StartYear Then

            ActiveCell.Value = TestValue(i)

            ActiveCell.Offset(1, 0).Select

        End If

    Next i

       

    StartYear = StartYear + 1

    

    RowNumber = ActiveCell.Row

    ColumnNumber = ActiveCell.Column

    

    ActiveSheet.Cells(155, ActiveCell.Column + 1).Select



Next k

   

Sheets("Long-Term Trends").Range("n3").Select

    

 '''''''''''''''''''''''''''''''''''''

 

Sheets("Long-Term Trends").Range("ae16").Select 'find calculated value of average (many digets)

LongTermAverage = Round(ActiveCell.Value, 2)    'round

If ActiveCell.Value < 1 Then LongTermAverage = Round(ActiveCell.Value, 3)   'round

Sheets("Long-Term Trends").Range("ae8").Select: ActiveCell.Value = "Average " + Parameter + " = " + CStr(LongTermAverage) + "  " 'create chart title



Sheets("Long-Term Trends").Range("n3").Select



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

Worksheets("Long-Term Trends").Activate

    Sheets("Long-Term Trends").Select

    TextBox1.Visible = False

    CommandButton4.Caption = "Open"

End Sub



