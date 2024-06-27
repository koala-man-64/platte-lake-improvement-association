Attribute VB_Name = "Sheet10"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Attribute VB_Control = "TextBox1, 1708, 1, MSForms, TextBox"
Attribute VB_Control = "CommandButton3, 1707, 2, MSForms, CommandButton"
Attribute VB_Control = "CommandButton1, 35, 3, MSForms, CommandButton"
Attribute VB_Control = "CommandButton2, 36, 4, MSForms, CommandButton"
Private Sub CommandButton3_Click()
If CommandButton3.Caption = "Open" Then
        CommandButton3.Caption = "Close"
        TextBox1.Visible = True
    Else
        CommandButton3.Caption = "Open"
        TextBox1.Visible = False
    End If
End Sub
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
If ActiveCell.Address <> "$S$11" And ActiveCell.Address <> "$S$12" And ActiveCell.Address <> "$S$13" And ActiveCell.Address <> "$S$14" And ActiveCell.Address <> "$S$15" And ActiveCell.Address <> "$S$16" And ActiveCell.Address <> "$S$17" And ActiveCell.Address <> "$S$18" And ActiveCell.Address <> "$S$19" And ActiveCell.Address <> "$S$20" And ActiveCell.Address <> "$S$21" And ActiveCell.Address <> "$S$22" And ActiveCell.Address <> "$S$23" And ActiveCell.Address <> "$S$24" And ActiveCell.Address <> "$S$25" Then Exit Sub
If ActiveCell.Value = "" Then Exit Sub
Application.ScreenUpdating = False
Range("s11:s25").Interior.ColorIndex = 2
ActiveCell.Interior.ColorIndex = 15
Dim TestDate As Date, SelectedYear As Integer
Dim DataCount As Integer, Parameter1 As String, Parameter2 As String, Parameter As String
Dim PlotValue(100) As Double
Dim PlotDate(100) As Date
Dim i As Integer, jj As Integer, k As Integer, j As Integer
Dim SelectedDate As Date
SelectedDate = ActiveCell.Value
Sheets("Near-Shore").Select
Sheets("Near-Shore").Range("b38").Select: DataCount = ActiveCell.Value
Sheets("Near-Shore").Range("i6").Select: SelectedYear = ActiveCell.Value
Sheets("Near-Shore").Range("k5").Select: Parameter1 = ActiveCell.Value
Sheets("Near-Shore").Range("k6").Select: Parameter2 = ActiveCell.Value
    For kk = 1 To 2     'kk because 2 charts
    If kk = 1 Then
        Parameter = Parameter1
        Sheets("Near-Shore").Range("aj4").Select: ActiveCell.Value = SelectedDate                           'define chart title date
    End If
    If kk = 2 Then
        Parameter = Parameter2
        Sheets("Near-Shore").Range("aj9").Select: ActiveCell.Value = SelectedDate                           'define chart title date
    End If
    k = 1                                      'k is the number of locations = 9 of the selected year
    Sheets("Near-Shore").Range("b40").Select   'date column
    For j = 1 To DataCount      'begin the scan through all the data
            TestDate = ActiveCell.Value     'select the first date record
            If TestDate = SelectedDate Then 'check to verify that record is for the selected fset
                    If Parameter = "Area" Then ActiveCell.Offset(0, 2).Select
                    If Parameter = "Rating" Then ActiveCell.Offset(0, 3).Select
                    If Parameter = "Temperature" Then ActiveCell.Offset(0, 4).Select
                    If Parameter = "Conductivity" Then ActiveCell.Offset(0, 5).Select
                        PlotValue(k) = ActiveCell.Value    'select the first date parameter value
                    If Parameter = "Area" Then ActiveCell.Offset(0, -2).Select
                    If Parameter = "Rating" Then ActiveCell.Offset(0, -3).Select
                    If Parameter = "Temperature" Then ActiveCell.Offset(0, -4).Select
                    If Parameter = "Conductivity" Then ActiveCell.Offset(0, -5).Select
                    k = k + 1                   'increment the count
            End If
     ActiveCell.Offset(1, 0).Select             'next record
    Next j
        If kk = 1 Then Sheets("Near-Shore").Range("ao2").Select              'print location based on site number
        If kk = 2 Then Sheets("Near-Shore").Range("ao3").Select              'print location based on site number
        For i = 1 To k - 1                          'this puts the values in place for the selected date fron the range
            ActiveCell.Value = PlotValue(i)
            ActiveCell.Offset(0, 1).Select
        Next i
    Next kk
    Sheets("Near-Shore").Range("g14").Select
End Sub
Private Sub worksheet_Activate()
Worksheets("Near-Shore").Activate
    Sheets("Near-Shore").Select
    TextBox1.Visible = False
    CommandButton3.Caption = "Open"
End Sub
Private Sub CommandButton1_Click()
Dim TestDate As Date, SelectedYear As Integer
Dim DataCount As Integer
Dim i As Integer, jj As Integer, k As Integer, j As Integer
Dim Parameter1 As String, Parameter2 As String, Parameter As String
Dim PlotDate(100) As Date, PlotValue(100) As Double
Dim SelectedDate As Date
Application.ScreenUpdating = False
Sheets("Near-Shore").Range("ah11:ah25").ClearContents   'clear labels
Sheets("Near-Shore").Range("s11:s25").ClearContents     'clear date list
Sheets("Near-Shore").Select
Sheets("Near-Shore").Range("b38").Select: DataCount = ActiveCell.Value
Sheets("Near-Shore").Range("i6").Select: SelectedYear = ActiveCell.Value
Sheets("Near-Shore").Range("k5").Select: Parameter1 = ActiveCell.Value
Sheets("Near-Shore").Range("k6").Select: Parameter2 = ActiveCell.Value
ActiveSheet.ChartObjects("Chart 6").Activate            'Right click chart and Chart Window
ActiveChart.ChartArea.Select
If Parameter1 = "Area" Then
    ActiveChart.Axes(xlValue).Select
    ActiveChart.Axes(xlValue).MinimumScale = 0
    ActiveChart.Axes(xlValue).MaximumScale = 1500
    ActiveChart.Axes(xlValue).MajorUnit = 300
    Selection.TickLabels.NumberFormat = "#,##0"
    Sheets("Near-Shore").Range("ai6").Select: ActiveCell.Value = Parameter1
End If
If Parameter1 = "Rating" Then
    ActiveChart.Axes(xlValue).Select
    ActiveChart.Axes(xlValue).MinimumScale = 0
    ActiveChart.Axes(xlValue).MaximumScale = 5
    ActiveChart.Axes(xlValue).MajorUnit = 1
    Selection.TickLabels.NumberFormat = "#,##0"
    Sheets("Near-Shore").Range("ai6").Select: ActiveCell.Value = Parameter1
End If
If Parameter1 = "Temperature" Then
    ActiveChart.Axes(xlValue).Select
    ActiveChart.Axes(xlValue).MinimumScale = 40
    ActiveChart.Axes(xlValue).MaximumScale = 90
    ActiveChart.Axes(xlValue).MajorUnit = 10
    Selection.TickLabels.NumberFormat = "#,##0"
    Sheets("Near-Shore").Range("ai6").Select: ActiveCell.Value = Parameter1
End If
If Parameter1 = "Conductivity" Then
    ActiveChart.Axes(xlValue).Select
    ActiveChart.Axes(xlValue).MinimumScale = 250
    ActiveChart.Axes(xlValue).MaximumScale = 400
    ActiveChart.Axes(xlValue).MajorUnit = 50
    Selection.TickLabels.NumberFormat = "#,##0"
    Sheets("Near-Shore").Range("ai6").Select: ActiveCell.Value = Parameter1
End If
ActiveSheet.ChartObjects("Chart 8").Activate            'Right click chart and Chart Window
ActiveChart.ChartArea.Select
If Parameter2 = "Area" Then
    ActiveChart.Axes(xlValue).Select
    ActiveChart.Axes(xlValue).MinimumScale = 0
    ActiveChart.Axes(xlValue).MaximumScale = 1500
    ActiveChart.Axes(xlValue).MajorUnit = 300
    Selection.TickLabels.NumberFormat = "#,##0"
    Sheets("Near-Shore").Range("ai16").Select: ActiveCell.Value = Parameter2
End If
If Parameter2 = "Rating" Then
    ActiveChart.Axes(xlValue).Select
    ActiveChart.Axes(xlValue).MinimumScale = 0
    ActiveChart.Axes(xlValue).MaximumScale = 5
    ActiveChart.Axes(xlValue).MajorUnit = 1
    Selection.TickLabels.NumberFormat = "#,##0"
    Sheets("Near-Shore").Range("ai16").Select: ActiveCell.Value = Parameter2
End If
If Parameter2 = "Temperature" Then
    ActiveChart.Axes(xlValue).Select
    ActiveChart.Axes(xlValue).MinimumScale = 40
    ActiveChart.Axes(xlValue).MaximumScale = 90
    ActiveChart.Axes(xlValue).MajorUnit = 10
    Selection.TickLabels.NumberFormat = "#,##0"
    Sheets("Near-Shore").Range("ai16").Select: ActiveCell.Value = Parameter2
End If
If Parameter2 = "Conductivity" Then
    ActiveChart.Axes(xlValue).Select
    ActiveChart.Axes(xlValue).MinimumScale = 250
    ActiveChart.Axes(xlValue).MaximumScale = 400
    ActiveChart.Axes(xlValue).MajorUnit = 50
    Selection.TickLabels.NumberFormat = "#,##0"
    Sheets("Near-Shore").Range("ai16").Select: ActiveCell.Value = Parameter2
End If
Sheets("Near-Shore").Range("b40").Select                    ''''''''''''''''''''''''''''''
jj = 1
For i = 1 To DataCount
    If Year(ActiveCell.Value) = SelectedYear Then
        TestDate = ActiveCell.Value
        If TestDate <> PlotDate(jj - 1) Then
             PlotDate(jj) = TestDate
             jj = jj + 1
        End If
    End If
    ActiveCell.Offset(1, 0).Select
Next i
SelectedDate = PlotDate(1)
Sheets("Near-Shore").Range("s11").Select        '''''''''''''''''''''''''''''''
For i = 1 To jj - 1                                'populate date list in coumn S
     ActiveCell.Value = PlotDate(i)
     ActiveCell.Offset(1, 0).Select
Next i
For kk = 1 To 2     'kk because 2 charts
If kk = 1 Then
    Parameter = Parameter1
    Sheets("Near-Shore").Range("aj4").Select: ActiveCell.Value = SelectedDate                           'define chart title date
    Sheets("Near-Shore").Range("aj5").Select: ActiveCell.Value = "Periphyton  " + Parameter + "  "      'define chart parameter
End If
If kk = 2 Then
    Parameter = Parameter2
    Sheets("Near-Shore").Range("aj9").Select: ActiveCell.Value = SelectedDate                           'define chart title date
    Sheets("Near-Shore").Range("aj10").Select: ActiveCell.Value = "Periphyton  " + Parameter + "  "      'define chart parameter
End If
k = 1                                      'k is the number of locations = 9 of the selected year
Sheets("Near-Shore").Range("b40").Select   'date column
For j = 1 To DataCount      'begin the scan through all the data
       TestDate = ActiveCell.Value     'select the first date record
       If TestDate = SelectedDate Then 'check to verify that record is for the selected fset
               If Parameter = "Area" Then ActiveCell.Offset(0, 2).Select
               If Parameter = "Rating" Then ActiveCell.Offset(0, 3).Select
               If Parameter = "Temperature" Then ActiveCell.Offset(0, 4).Select
               If Parameter = "Conductivity" Then ActiveCell.Offset(0, 5).Select
               PlotValue(k) = ActiveCell.Value    'select the first date parameter value
               If Parameter = "Area" Then ActiveCell.Offset(0, -2).Select
               If Parameter = "Rating" Then ActiveCell.Offset(0, -3).Select
               If Parameter = "Temperature" Then ActiveCell.Offset(0, -4).Select
               If Parameter = "Conductivity" Then ActiveCell.Offset(0, -5).Select
               k = k + 1                   'increment the count
       End If
 ActiveCell.Offset(1, 0).Select             'next record
 Next j
        If kk = 1 Then Sheets("Near-Shore").Range("ao2").Select              'print location based on site number
        If kk = 2 Then Sheets("Near-Shore").Range("ao3").Select              'print location based on site number
        For i = 1 To k - 1                          'this puts the values in place for the selected date fron the range
            ActiveCell.Value = PlotValue(i)
            ActiveCell.Offset(0, 1).Select
        Next i
    Next kk
    Sheets("Near-Shore").Range("g14").Select
Sheets("Near-Shore").Range("s11").Select
End Sub
Private Sub CommandButton2_Click()
 Sheets("Main Menu").Select
    Sheets("Main Menu").Range("g11").Select
End Sub