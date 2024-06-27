Attribute VB_Name = "Sheet4"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Attribute VB_Control = "CommandButton1, 1, 0, MSForms, CommandButton"
Attribute VB_Control = "CommandButton2, 2, 1, MSForms, CommandButton"
Attribute VB_Control = "CommandButton3, 3, 2, MSForms, CommandButton"
Attribute VB_Control = "CommandButton4, 4, 3, MSForms, CommandButton"
Attribute VB_Control = "TextBox1, 5, 4, MSForms, TextBox"
Private Sub CommandButton4_Click()
Dim TestLat(500) As Double
Dim TestLong(500) As Double
Sheets("Shoreline").Range("ak4:aL9000").ClearContents           'LEFT CHART UPDATE XXXXXXXXXXXXXXXXXX
Application.ScreenUpdating = False
Sheets("Shoreline").Select
Sheets("Shoreline").Range("r13").Select: TopWeight = ActiveCell.Value
Sheets("Shoreline").Range("r14").Select: TopCond = ActiveCell.Value
    Sheets("Shoreline").Range("f10").Select
    PlotYear = ActiveCell.Value              'year
    ActiveCell.Offset(1, 0).Select
    Select1 = ActiveCell.Value               'cause
    ActiveCell.Offset(1, 0).Select
    Select2 = ActiveCell.Value               'substrate
    ActiveCell.Offset(1, 0).Select
    Select3 = ActiveCell.Value               'density  high or all
    ActiveCell.Offset(1, 0).Select
    Select4 = ActiveCell.Value               'cond
    If Select1 = "Septic" Then Select1 = "S"
    If Select1 = "Fertilizer" Then Select1 = "F"
    If Select1 = "Outfall" Then Select1 = "O"
    If Select1 = "Uncertain" Then Select1 = "U"
    If Select2 = "Rocks" Then Select2 = "R"
    If Select2 = "Pebbles" Then Select2 = "P"
    If Select2 = "Sand" Then Select2 = "S"
    If Select2 = "Sea Wall" Then Select2 = "W"
    k = 0
    TotalWeight = 0
    Sheets("Shoreline").Range("b55").Select: StoredDataCount = ActiveCell.Value
    Sheets("Shoreline").Range("b57").Select
For j = 1 To StoredDataCount
        TestDate = ActiveCell.Value             'read date
        ActiveCell.Offset(0, 1).Select
        Lat = ActiveCell.Value                  'read lat
        ActiveCell.Offset(0, 1).Select
        Lon = ActiveCell.Value                  'read long
        ActiveCell.Offset(0, 5).Select
        Cause = ActiveCell.Value                'read cause
        ActiveCell.Offset(0, 1).Select
        Substrate = ActiveCell.Value            'read substrate
        ActiveCell.Offset(0, 1).Select
        cond = ActiveCell.Value                 'read cond
        ActiveCell.Offset(0, 14).Select
        Weight = ActiveCell.Value               'read Weight
     If (Select3 = "High Wt" And Weight > TopWeight) Or Select3 = "Any" Then      'if you select high and the value is high then TRUE
            WeightLogical = True                                                   'if you select ANY the value is TRUE regardless of the value of Weight
          Else
            WeightLogical = False
        End If
        If (Select4 = "High Cond" And cond > TopCond) Or Select4 = "Any" Then
            CondLogical = True
          Else
            CondLogical = False
        End If
        If Year(TestDate) = PlotYear And (Cause = Select1 Or Select1 = "Any") And (Substrate = Select2 Or Select2 = "Any") And WeightLogical = True And CondLogical = True Then
            k = k + 1
            TestLat(k) = Lat
            TestLong(k) = Lon
            TotalWeight = TotalWeight + Weight / 2000
        End If
        ActiveCell.Offset(1, -23).Select
Next j
 Sheets("Shoreline").Range("ak4").Select       'copy stuff to save
    For i = 1 To k
        ActiveCell.Value = Round((TestLong(i)), 4)
        ActiveCell.Offset(0, 1).Select
        ActiveCell.Value = Round((TestLat(i)), 4)
        ActiveCell.Offset(1, -1).Select
    Next i
ActiveSheet.ChartObjects("Chart 7").Activate           'Right click chart and Chart Window
ActiveChart.Axes(xlCategory).MajorUnit = 0.02
ActiveChart.ChartArea.Select
With ActiveChart.Axes(xlCategory)
     .MinimumScale = -86.26
     .MaximumScale = -86.06
End With
Sheets("Shoreline").Range("f16").Select: ActiveCell.Value = TotalWeight
Sheets("Shoreline").Range("f10").Select
End Sub
Private Sub CommandButton1_Click()
Dim TestLat(500) As Double
Dim TestLong(500) As Double
Sheets("Shoreline").Range("aN4:aO9000").ClearContents           'RIGHT CHART UPDATE XXXXXXXXXXXXXXXXXX
Application.ScreenUpdating = False
Sheets("Shoreline").Select
Sheets("Shoreline").Range("r13").Select: TopWeight = ActiveCell.Value
Sheets("Shoreline").Range("r14").Select: TopCond = ActiveCell.Value
    Sheets("Shoreline").Range("L10").Select
    PlotYear = ActiveCell.Value              'year
    ActiveCell.Offset(1, 0).Select
    Select1 = ActiveCell.Value               'cause
    ActiveCell.Offset(1, 0).Select
    Select2 = ActiveCell.Value               'substrate
    ActiveCell.Offset(1, 0).Select
    Select3 = ActiveCell.Value               'density  high or all
    ActiveCell.Offset(1, 0).Select
    Select4 = ActiveCell.Value               'cond
    If Select1 = "Septic" Then Select1 = "S"
    If Select1 = "Fertilizer" Then Select1 = "F"
    If Select1 = "Outfall" Then Select1 = "O"
    If Select1 = "Uncertain" Then Select1 = "U"
    If Select2 = "Rocks" Then Select2 = "R"
    If Select2 = "Pebbles" Then Select2 = "P"
    If Select2 = "Sand" Then Select2 = "S"
    If Select2 = "Sea Wall" Then Select2 = "W"
    k = 0
    TotalWeight = 0
    Sheets("Shoreline").Range("b55").Select: StoredDataCount = ActiveCell.Value
    Sheets("Shoreline").Range("b57").Select
For j = 1 To StoredDataCount
        TestDate = ActiveCell.Value             'read date
        ActiveCell.Offset(0, 1).Select
        Lat = ActiveCell.Value                  'read lat
        ActiveCell.Offset(0, 1).Select
        Lon = ActiveCell.Value                  'read long
        ActiveCell.Offset(0, 5).Select
        Cause = ActiveCell.Value                'read cause
        ActiveCell.Offset(0, 1).Select
        Substrate = ActiveCell.Value            'read substrate
        ActiveCell.Offset(0, 1).Select
        cond = ActiveCell.Value                 'read cond
        ActiveCell.Offset(0, 14).Select
        Weight = ActiveCell.Value               'read Weight
     If (Select3 = "High Wt" And Weight > TopWeight) Or Select3 = "Any" Then      'if you select high and the value is high then TRUE
            WeightLogical = True                                                   'if you select ANY the value is TRUE regardless of the value of Weight
          Else
            WeightLogical = False
        End If
        If (Select4 = "High Cond" And cond > TopCond) Or Select4 = "Any" Then
            CondLogical = True
          Else
            CondLogical = False
        End If
        If Year(TestDate) = PlotYear And (Cause = Select1 Or Select1 = "Any") And (Substrate = Select2 Or Select2 = "Any") And WeightLogical = True And CondLogical = True Then
            k = k + 1
            TestLat(k) = Lat
            TestLong(k) = Lon
            TotalWeight = TotalWeight + Weight / 2000
        End If
        ActiveCell.Offset(1, -23).Select
Next j
 Sheets("Shoreline").Range("an4").Select       'copy stuff to save
    For i = 1 To k
        ActiveCell.Value = Round((TestLong(i)), 4)
        ActiveCell.Offset(0, 1).Select
        ActiveCell.Value = Round((TestLat(i)), 4)
        ActiveCell.Offset(1, -1).Select
    Next i
ActiveSheet.ChartObjects("Chart 8").Activate           'Right click chart and Chart Window
ActiveChart.Axes(xlCategory).MajorUnit = 0.02
ActiveChart.ChartArea.Select
With ActiveChart.Axes(xlCategory)
     .MinimumScale = -86.26
     .MaximumScale = -86.06
End With
Sheets("Shoreline").Range("L16").Select: ActiveCell.Value = TotalWeight
End Sub
Private Sub worksheet_Activate()
Worksheets("Shoreline").Activate
    Sheets("Shoreline").Select
    TextBox1.Visible = False
    CommandButton3.Caption = "Open"
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