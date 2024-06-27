Attribute VB_Name = "Sheet19"

Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"

Attribute VB_GlobalNameSpace = False

Attribute VB_Creatable = False

Attribute VB_PredeclaredId = True

Attribute VB_Exposed = True

Attribute VB_TemplateDerived = False

Attribute VB_Customizable = True

Attribute VB_Control = "TextBox1, 4, 0, MSForms, TextBox"

Attribute VB_Control = "CommandButton3, 3, 1, MSForms, CommandButton"

Attribute VB_Control = "CommandButton2, 2, 2, MSForms, CommandButton"

Attribute VB_Control = "CommandButton1, 1, 3, MSForms, CommandButton"

Private Sub CommandButton1_Click()

Dim Flow(500) As Double, WeightedFlow(500) As Double, WetFlow(500) As Double, Rain(500) As Double

Dim EventFactor As Double, SumWetflow As Double, AverageDryFlow As Double, AverageFlow As Double, TotalRain As Double, AverageWetFlow

Dim i As Integer, Spikes As Integer, WetFlowCount As Integer, FlowYear As Integer

Dim Data1Count As Integer

Dim TestDate As Date, TestValue As Double

Dim PlotDate(4000) As Date, PlotValue(4000) As Double

Dim Site As String



Application.ScreenUpdating = False  'FIND EVENTS



Sheets("Flow & Rain & TP Comparison").Select

Sheets("Flow & Rain & TP Comparison").Range("az40:bb1000").ClearContents

Sheets("Flow & Rain & TP Comparison").Range("bc40:bd1000").ClearContents



Sheets("Flow & Rain & TP Comparison").Range("j3").Select

FlowYear = ActiveCell.Value

ActiveCell.Offset(1, 0).Select

Site = ActiveCell.Value

Sheets("Flow & Rain & TP Comparison").Range("n" + CStr(FlowYear - 2010 + 10)).Select

EventFactor = ActiveCell.Value



If EventFactor <= 0 Then          'check that data are available for selected range

   response% = MsgBox("Enter the Event Sensivity for " & CStr(FlowYear) & " in Column N." & " The Value Must be Greater than 0.", 64)

   Exit Sub

End If

   

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

 

Sheets("Flow & Rain Data").Select



SelectColumn = 25 + (FlowYear - 2010)            '25 is column Y for 2010 for RAIN



ActiveSheet.Cells(11, SelectColumn).Select      'find total rain for selected year

If ActiveCell.Value < 0.1 Or ActiveCell.Value = "" Then                'check that data are available for selected range

    response% = MsgBox("Rain Data for " & CStr(FlowYear) & " have Not been Entered", 64)

    Exit Sub

End If



ActiveSheet.Cells(13, SelectColumn).Select  'first data

For j = 1 To 366

     Rain(j) = ActiveCell.Value             'get daily rain

    ActiveCell.Offset(1, 0).Select

Next j



SelectColumn = 25 + (FlowYear - 1990)            '25 is column Y for 1990 for FLOW



ActiveSheet.Cells(381, SelectColumn).Select     'find average flow for selected year

If ActiveCell.Value < 0.1 Or ActiveCell.Value = "" Then                'check that data are available for selected range

    response% = MsgBox("Flow Data for " & CStr(FlowYear) & " have Not been Entered", 64)

    Exit Sub

End If



ActiveSheet.Cells(383, SelectColumn).Select 'first data

For j = 1 To 366

     Flow(j) = ActiveCell.Value             'get daily flow

    ActiveCell.Offset(1, 0).Select

Next j



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



Sheets("Flow & Rain & TP Comparison").Select                    'print daily flow

Sheets("Flow & Rain & TP Comparison").Range("az40").Select

For i = 1 To 366

    ActiveCell.Value = Flow(i)

    ActiveCell.Offset(1, 0).Select

Next i



Sheets("Flow & Rain & TP Comparison").Range("az39").Select: AverageFlow = ActiveCell.Value

Sheets("Flow & Rain & TP Comparison").Range("r" + CStr(FlowYear - 2010 + 10)).Select: ActiveCell.Value = AverageFlow



Sheets("Flow & Rain & TP Comparison").Range("bb40").Select          'print daily rain

For i = 1 To 366

    ActiveCell.Value = Rain(i)

    ActiveCell.Offset(1, 0).Select

Next i



WeightedFlow(1) = Flow(1)

WeightedFlow(2) = Flow(2)

WeightedFlow(3) = (Flow(1) + Flow(2)) / 2

WeightedFlow(4) = (Flow(1) + Flow(2) + Flow(3)) / 3



For i = 5 To 365

    WeightedFlow(i) = (Flow(i - 4) + Flow(i - 3) + Flow(i - 2) + Flow(i - 1)) / 4

Next i



Spikes = 0

SumWetflow = 0

WetFlowCount = 0



For i = 1 To 365

    If Flow(i) > WeightedFlow(i) * (1 + EventFactor / 100) Then

        

        WetFlowCount = WetFlowCount + 1

        Spikes = Spikes + 1

        SumWetflow = SumWetflow + Flow(i)

        WetFlow(i) = Flow(i)

        

    End If

Next i

   

Sheets("Flow & Rain & TP Comparison").Range("ba40").Select      'print few event flows

For i = 1 To 365

    If WetFlow(i) <> 0 Then ActiveCell.Value = WetFlow(i)

    ActiveCell.Offset(1, 0).Select

Next i



Sheets("Flow & Rain & TP Comparison").Range("o" + CStr(FlowYear - 2010 + 10)).Select

ActiveCell.Value = Spikes

ActiveCell.Offset(0, 1).Select



AverageWetFlow = SumWetflow / Spikes

ActiveCell.Value = AverageWetFlow

ActiveCell.Offset(0, 1).Select



AverageDryFlow = (365 * AverageFlow - AverageWetFlow * Spikes) / (365 - Spikes)

ActiveCell.Value = AverageDryFlow



Sheets("Flow & Rain & TP Comparison").Range("ba3").Select           'define chart labels

ActiveCell.Value = "Flow  " + CStr(FlowYear) + "   Events = " + CStr(Spikes)

Sheets("Flow & Rain & TP Comparison").Range("bb3").Select           'define chart labels

ActiveCell.Value = "Rain " + CStr(FlowYear)



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



If Site = "Lake TP" Then



    Sheets("Lake Chemistry").Select

    Sheets("Lake Chemistry").Range("f37").Select: Data1Count = ActiveCell.Value

    ActiveCell.Offset(2, -4).Select

    

    k = 1

    For j = 1 To Data1Count      'begin the scan through all the data

                        

                TestDate = ActiveCell.Value     'select the first date record

                ActiveCell.Offset(0, 4).Select

                TestValue = ActiveCell.Value    'select the first date parameter value

               

                If Year(TestDate) = FlowYear Then

                    PlotDate(k) = TestDate

                    PlotValue(k) = TestValue

                k = k + 1                   'increment the count

                End If

                

                ActiveCell.Offset(1, -4).Select     'go down one and back to the date column

    Next j

    

End If





If Site <> "Lake TP" Then



    Sheets("Stream Chemistry").Select

    If Site = "Stone" Then Sheets("Stream Chemistry").Range("c38").Select: Data1Count = ActiveCell.Value

    If Site = "Vet's" Then Sheets("Stream Chemistry").Range("f38").Select: Data1Count = ActiveCell.Value

    If Site = "Haze" Then Sheets("Stream Chemistry").Range("i38").Select: Data1Count = ActiveCell.Value

    If Site = "Carter" Then Sheets("Stream Chemistry").Range("l38").Select: Data1Count = ActiveCell.Value

    If Site = "Pioneer" Then Sheets("Stream Chemistry").Range("o38").Select: Data1Count = ActiveCell.Value

    If Site = "USGS" Then Sheets("Stream Chemistry").Range("r38").Select: Data1Count = ActiveCell.Value

    If Site = "NB Ind Hill" Then Sheets("Stream Chemistry").Range("u38").Select: Data1Count = ActiveCell.Value

    If Site = "NB Dead" Then Sheets("Stream Chemistry").Range("x38").Select: Data1Count = ActiveCell.Value

    If Site = "NB Hooker" Then Sheets("Stream Chemistry").Range("aa38").Select: Data1Count = ActiveCell.Value

    If Site = "M22" Then Sheets("Stream Chemistry").Range("ad38").Select: Data1Count = ActiveCell.Value

    If Site = "BC Old Res" Then Sheets("Stream Chemistry").Range("ag38").Select: Data1Count = ActiveCell.Value

    If Site = "Collision" Then Sheets("Stream Chemistry").Range("aj38").Select: Data1Count = ActiveCell.Value

    

    ActiveCell.Offset(2, -1).Select

    

    k = 0

    For j = 1 To Data1Count      'begin the scan through all the data

                        

                TestDate = ActiveCell.Value     'select the first date record

                ActiveCell.Offset(0, 1).Select

                TestValue = ActiveCell.Value    'select the first date parameter value

                

                If Year(TestDate) = FlowYear Then

                    PlotDate(k) = TestDate

                    PlotValue(k) = TestValue

                k = k + 1                   'increment the count

                End If

                

                ActiveCell.Offset(1, -1).Select     'go down one and back to the date column

    Next j



End If



Sheets("Flow & Rain & TP Comparison").Select

Sheets("Flow & Rain & TP Comparison").Range("bc40").Select     'print range for plot

For i = 1 To k - 1                                                                      'need to substract one because of k=k+1 above

    dayofyear = DateDiff("d", CDate("1/1/" & Year(PlotDate(i))), PlotDate(i)) + 1

    ActiveCell.Value = dayofyear

    ActiveCell.Offset(0, 1).Select

    ActiveCell.Value = PlotValue(i)

    ActiveCell.Offset(1, -1).Select

Next i



Sheets("Flow & Rain & TP Comparison").Range("bd38").Select

If ActiveCell.Value <> "" Then meanTP = Round(ActiveCell.Value, 2)



Sheets("Flow & Rain & TP Comparison").Range("bc3").Select           'define chart labels

ActiveCell.Value = Site & " Avg = " + CStr(meanTP) + "  " + CStr(FlowYear)





Sheets("Flow & Rain & TP Comparison").Range("j3").Select



End Sub

Private Sub CommandButton2_Click()

    Sheets("Main Menu").Select

    Sheets("Main Menu").Range("g11").Select

End Sub

Private Sub CommandButton3_Click()

    If CommandButton3.Caption = "Open" Then

        CommandButton3.Caption = "Close"

        TextBox1.Visible = True

    Else

        CommandButton3.Caption = "Open"

        TextBox1.Visible = False

    End If

End Sub



Private Sub worksheet_Activate()

Worksheets("Flow & Rain & TP Comparison").Activate

    Sheets("Flow & Rain & TP Comparison").Select

    TextBox1.Visible = False

    CommandButton3.Caption = "Open"

End Sub









