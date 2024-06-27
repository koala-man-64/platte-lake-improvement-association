Attribute VB_Name = "Sheet2"

Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"

Attribute VB_GlobalNameSpace = False

Attribute VB_Creatable = False

Attribute VB_PredeclaredId = True

Attribute VB_Exposed = True

Attribute VB_TemplateDerived = False

Attribute VB_Customizable = True

Attribute VB_Control = "CommandButton3, 4, 1, MSForms, CommandButton"

Attribute VB_Control = "CommandButton2, 3, 2, MSForms, CommandButton"

Attribute VB_Control = "CommandButton1, 1, 3, MSForms, CommandButton"

Attribute VB_Control = "TextBox1, 6, 4, MSForms, TextBox"

Private Sub CommandButton2_Click()

If CommandButton2.Caption = "Open" Then

        CommandButton2.Caption = "Close"

        TextBox1.Visible = True

    Else

        CommandButton2.Caption = "Open"

        TextBox1.Visible = False

    End If

End Sub

Private Sub worksheet_Activate()

Worksheets("Wet Weather TP").Activate

    Sheets("Wet Weather TP").Select

    TextBox1.Visible = False

    CommandButton2.Caption = "Open"

End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)    'this code runs if anything is changed

Dim PlotValue(500) As Double, PlotTime(500) As Double, PlotDay(500) As Double, PlotFlow(500) As Double, PlotDate(500) As Date

Dim SelectedDate As Date, TestDate As Date

Dim SelectedSite As String

Dim DataCount As Integer



Application.ScreenUpdating = False



If ActiveCell.Address <> "$T$9" And ActiveCell.Address <> "$T$10" And ActiveCell.Address <> "$T$11" And ActiveCell.Address <> "$T$12" And ActiveCell.Address <> "$T$13" And ActiveCell.Address <> "$T$14" And ActiveCell.Address <> "$T$15" And ActiveCell.Address <> "$T$16" And ActiveCell.Address <> "$T$17" And ActiveCell.Address <> "$T$18" And ActiveCell.Address <> "$T$19" And ActiveCell.Address <> "$T$20" And ActiveCell.Address <> "$T$21" And ActiveCell.Address <> "$T$22" And ActiveCell.Address <> "$T$23" And ActiveCell.Address <> "$T$24" And ActiveCell.Address <> "$T$25" And ActiveCell.Address <> "$T$26" And ActiveCell.Address <> "$T$27" And ActiveCell.Address <> "$T$28" And ActiveCell.Address <> "$T$29" And ActiveCell.Address <> "$T$30" Then Exit Sub



If ActiveCell.Value = "" Then Exit Sub



Sheets("Wet Weather TP").Range("ai41:ai500").ClearContents

Sheets("Wet Weather TP").Range("aj41:ak500").ClearContents

    

    Range("T9:T30").Interior.ColorIndex = 2    'make all cells white again after selection

        

    ActiveCell.Interior.ColorIndex = 15

        

    SelectedDate = ActiveCell.Value

    Sheets("Wet Weather TP").Range("af7").Select:  ActiveCell.Value = SelectedDate

    

    Sheets("Wet Weather TP").Range("i6").Select

    SelectedSite = ActiveCell.Value

    

    If SelectedSite = "Stone" Then Sheets("Wet Weather TP").Range("c35").Select: DataCount = ActiveCell.Value

    If SelectedSite = "Brundage" Then Sheets("Wet Weather TP").Range("g35").Select: DataCount = ActiveCell.Value

    If SelectedSite = "USGS" Then Sheets("Wet Weather TP").Range("k35").Select: DataCount = ActiveCell.Value



    If SelectedSite = "Stone" Then Sheets("Wet Weather TP").Range("b41").Select

    If SelectedSite = "Brundage" Then Sheets("Wet Weather TP").Range("f41").Select

    If SelectedSite = "USGS" Then Sheets("Wet Weather TP").Range("j41").Select

    

k = 1

For j = 1 To DataCount



    TestDate = ActiveCell.Value

 

    If TestDate = SelectedDate Then

        ActiveCell.Offset(0, 1).Select

        PlotValue(k) = ActiveCell.Value

       

        ActiveCell.Offset(0, 1).Select

        PlotTime(k) = ActiveCell.Value

        k = k + 1

        ActiveCell.Offset(0, -2).Select

    End If



    ActiveCell.Offset(1, 0).Select

    

Next j

 

Sheets("Wet Weather TP").Range("aj41").Select

        

        For i = 1 To k - 1

          ActiveCell.Value = PlotTime(i)

          ActiveCell.Offset(0, 1).Select

          ActiveCell.Value = PlotValue(i)

          ActiveCell.Offset(1, -1).Select

        Next i

'End If

        

     If Year(SelectedDate) = 2003 Then Sheets("Wet Weather TP").Range("m41").Select     'prepare to print flows

     If Year(SelectedDate) = 2004 Then Sheets("Wet Weather TP").Range("q41").Select

     If Year(SelectedDate) = 2005 Then Sheets("Wet Weather TP").Range("u41").Select

     If Year(SelectedDate) = 2006 Then Sheets("Wet Weather TP").Range("y41").Select

     If Year(SelectedDate) = 2007 Then Sheets("Wet Weather TP").Range("ac41").Select

        

     For j = 1 To 365

     

        PlotDay(j) = ActiveCell.Value

        ActiveCell.Offset(0, 1).Select

        PlotDate(j) = ActiveCell.Value

        ActiveCell.Offset(0, 1).Select

        PlotFlow(j) = ActiveCell.Value

        ActiveCell.Offset(0, -2).Select

        

        If PlotDate(j) = SelectedDate Then

          Spikeday = PlotDay(j)

          SpikeDate = PlotDate(j)

          SpikeFlow = PlotFlow(j)

        End If

        

        ActiveCell.Offset(1, 0).Select

        

     Next j

     

       Sheets("Wet Weather TP").Range("ah41").Select  'print flows for selected year

       For j = 1 To 365

         ActiveCell.Value = PlotFlow(j)

         ActiveCell.Offset(1, 0).Select

       Next j

       

       Sheets("Wet Weather TP").Range("ag41").Select  'print spike flow

       For j = 1 To 365

         If ActiveCell.Value = Spikeday Then

            ActiveCell.Offset(0, 2).Select

            ActiveCell.Value = SpikeFlow

            Exit For

         End If

       ActiveCell.Offset(1, 0).Select

       Next j

       

       

     Sheets("Wet Weather TP").Range("ak38").Select  'find maxiumum

     MaxValue = ActiveCell.Value

     

     ActiveSheet.ChartObjects("Chart 9").Activate

     ActiveChart.ChartArea.Select

  

     If MaxValue < 100 Then

        With ActiveChart.Axes(xlValue)

            .MinimumScale = 0

            .MaximumScale = 100

            .MajorUnit = 20

        End With

     End If

       

     If MaxValue > 100 And MaxValue < 200 Then

        With ActiveChart.Axes(xlValue)

            .MinimumScale = 0

            .MaximumScale = 200

            .MajorUnit = 40

        End With

     End If

     

     If MaxValue > 200 And MaxValue < 300 Then

        With ActiveChart.Axes(xlValue)

            .MinimumScale = 0

            .MaximumScale = 300

            .MajorUnit = 50

        End With

     End If

     

     If MaxValue > 300 Then

        With ActiveChart.Axes(xlValue)

            .MinimumScale = 0

            .MaximumScale = 500

            .MajorUnit = 100

        End With

     End If

       

Sheets("Wet Weather TP").Range("i4").Select

        

End Sub

Private Sub CommandButton1_Click()

Dim TestDate As Date, SelectedYear As Integer, SelectedSite As String

Dim DataCount As Integer

Dim PlotValue(500) As Double

Dim PlotDate(500) As Date, PlotTime(500) As Double

Dim i As Integer, jj As Integer, k As Integer, j As Integer



Application.ScreenUpdating = False

Sheets("Wet Weather TP").Range("ak9:ak30").ClearContents



Sheets("Wet Weather TP").Select

Sheets("Wet Weather TP").Range("i4").Select: SelectedYear = ActiveCell.Value

Sheets("Wet Weather TP").Range("i6").Select: SelectedSite = ActiveCell.Value



If SelectedSite = "Stone" Then Sheets("Wet Weather TP").Range("c35").Select: DataCount = ActiveCell.Value

If SelectedSite = "Brundage" Then Sheets("Wet Weather TP").Range("g35").Select: DataCount = ActiveCell.Value

If SelectedSite = "USGS" Then Sheets("Wet Weather TP").Range("k35").Select: DataCount = ActiveCell.Value



Sheets("Wet Weather TP").Range("af4").Select: ActiveCell.Value = SelectedYear                        'define chart title date

Sheets("Wet Weather TP").Range("af5").Select: ActiveCell.Value = SelectedSite                        'define chart title date



ActiveSheet.ChartObjects("Chart 9").Activate            'Right click chart and Chart Window

    ActiveChart.ChartArea.Select

    ActiveChart.Axes(xlValue).Select

    ActiveChart.Axes(xlValue).MinimumScale = 0

    ActiveChart.Axes(xlValue).MaximumScale = 25

    ActiveChart.Axes(xlValue).MajorUnit = 5

    Selection.TickLabels.NumberFormat = "#,##0"



'Loop  through data to find dates for selected year

If SelectedSite = "Stone" Then Sheets("Wet Weather TP").Range("b41").Select

If SelectedSite = "Brundage" Then Sheets("Wet Weather TP").Range("f41").Select

If SelectedSite = "USGS" Then Sheets("Wet Weather TP").Range("j41").Select



k = 1



For i = 1 To DataCount              'begin the scan through all the data

  

  TestDate = ActiveCell.Value       'select the first date record

  

  If Year(TestDate) = SelectedYear And TestDate <> LastTestDate Then      'note PlotDate(0) is small

      PlotDate(k) = TestDate

      LastTestDate = TestDate

      k = k + 1

  End If

  

  ActiveCell.Offset(1, 0).Select

    

Next i



Sheets("Wet Weather TP").Range("ak9").Select     'print range for plot

    For j = 1 To k - 1

        ActiveCell.Value = PlotDate(j)

        ActiveCell.Offset(1, 0).Select

    Next j

 

Sheets("Wet Weather TP").Range("t9").Select



End Sub

Private Sub CommandButton3_Click()

    Sheets("Main Menu").Select

    Sheets("Main Menu").Range("g11").Select

End Sub



