Attribute VB_Name = "Sheet15"
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
Private Sub CommandButton1_Click()
Dim TestValue(100) As Variant, PlotValue(100) As Variant
Dim YLabel As String, Site As String
'Annual Averages  Charts 1 & 2
Application.ScreenUpdating = False
Sheets("Annual Averages").Select
Sheets("Annual Averages").Range("ap5:ap57").ClearContents 'clear plot data from previous run         'clear previous results
Sheets("Annual Averages").Range("aq5:aq57").ClearContents
Sheets("Annual Averages").Range("i3").Select: Parameter1 = ActiveCell.Value
Sheets("Annual Averages").Range("i4").Select: Parameter2 = ActiveCell.Value
For k = 1 To 2
    If k = 1 Then Parameter = Parameter1
    If k = 2 Then Parameter = Parameter2
     If Parameter = "Days>8" Then YLabel = "Number of Days": XMax = 300: XUnit = 50: Range("d48").Select
     If Parameter = "Loss Rate" Then YLabel = "meters/year": XMax = 40: XUnit = 10: Range("f48").Select
     If Parameter = "USGS Flow" Then YLabel = "cfs": XMax = 300: XUnit = 50: Range("z48").Select
     If Parameter = "Rain Inch" Then YLabel = "Inches": XMax = 60: XUnit = 10: Range("aa48").Select
     If Parameter = "BC InFlow" Then YLabel = "mgd": XMax = 15: XUnit = 3: Range("p48").Select
     If Parameter = "PRSFH OutFlow" Then YLabel = "mgd": XMax = 15: XUnit = 3: Range("s48").Select
     If Parameter = "PRSFH Load" Then YLabel = "Pounds": XMax = 600: XUnit = 100: Range("u48").Select
     If Parameter = "BC Load" Then YLabel = "Pounds": XMax = 600: XUnit = 100: Range("r48").Select
     If Parameter = "Lost Fish" Then YLabel = "Pounds": XMax = 300: XUnit = 50: Range("v48").Select
     If Parameter = "Sed Rel" Then YLabel = "Pounds": XMax = 400: XUnit = 100: Range("e48").Select
     If Parameter = "Total Load" Then YLabel = "Pounds": XMax = 12000: XUnit = 3000: Range("w48").Select
     If Parameter = "Lower NP" Then YLabel = "Pounds": XMax = 12000: XUnit = 3000: Range("x48").Select
     If Parameter = "Upper NP" Then YLabel = "Pounds": XMax = 12000: XUnit = 3000: Range("y48").Select
     If Parameter = "Rain Load" Then YLabel = "Pounds": XMax = 600: XUnit = 100: Range("ab48").Select
   If Parameter = "Lake TP" Then YLabel = "TP  mg/m3": XMax = 20: XUnit = 4: Range("c48").Select
   If Parameter = "Stone TP" Then YLabel = "TP  mg/m3": XMax = 32: XUnit = 4: Range("g48").Select
   If Parameter = "Carter TP" Then YLabel = "TP  mg/m3": XMax = 32: XUnit = 4: Range("h48").Select
   If Parameter = "Collision TP" Then YLabel = "TP  mg/m3": XMax = 32: XUnit = 4: Range("i48").Select
   If Parameter = "NB Dead TP" Then YLabel = "TP  mg/m3": XMax = 32: XUnit = 4: Range("j48").Select
   If Parameter = "Vet's TP" Then YLabel = "TP  mg/m3": XMax = 32: XUnit = 4: Range("k48").Select
   If Parameter = "Pioneer TP" Then YLabel = "TP  mg/m3": XMax = 32: XUnit = 4: Range("L48").Select
   If Parameter = "USGS TP" Then YLabel = "TP  mg/m3": XMax = 32: XUnit = 4: Range("m48").Select
   If Parameter = "Haze TP" Then YLabel = "TP  mg/m3": XMax = 32: XUnit = 4: Range("n48").Select
   If Parameter = "NB Ind Hill TP" Then YLabel = "TP  mg/m3": XMax = 32: XUnit = 4: Range("o48").Select
   If Parameter = "PRSFH TP" Then YLabel = "TP  mg/m3": XMax = 32: XUnit = 4: Range("t48").Select
   If Parameter = "BC TP" Then YLabel = "TP  mg/m3": XMax = 32: XUnit = 4: Range("q48").Select
   For j = 1 To 40     ' read'''''''''''''''''''''''''''''''''''''''''''''''''
      TestValue(j) = ActiveCell.Value
      If TestValue(j) = 0 Or TestValue(j) = "" Then Exit For
      ActiveCell.Offset(1, 0).Select
   Next j
   If k = 1 Then Range("ap7").Select
   If k = 2 Then Range("aq7").Select
   For i = 1 To 40     'print''''''''''''''''''''''''''''''''''''''''''''''''
      ActiveCell.Value = TestValue(i)
      ActiveCell.Offset(1, 0).Select
   Next i
   If k = 1 Then Sheets("Annual Averages").Range("ap5").Select: ActiveCell.Value = YLabel:  ActiveCell.Offset(1, 0).Select: ActiveCell.Value = "Annual Average  " & Parameter
   If k = 2 Then Sheets("Annual Averages").Range("aq5").Select: ActiveCell.Value = YLabel:  ActiveCell.Offset(1, 0).Select: ActiveCell.Value = "Annual Average  " & Parameter
If k = 1 Then
ActiveSheet.ChartObjects("Chart 8").Activate
        With ActiveChart.Axes(xlValue)
            .MinimumScale = 0
            .MaximumScale = XMax
            .MajorUnit = XUnit
        End With
End If
If k = 2 Then
ActiveSheet.ChartObjects("Chart 9").Activate
        With ActiveChart.Axes(xlValue)
            .MinimumScale = 0
            .MaximumScale = XMax
            .MajorUnit = XUnit
        End With
End If
Next k
Sheets("Annual Averages").Range("i3").Select
End Sub
Private Sub CommandButton2_Click()
Dim TestValue(100) As Variant, PlotValue(100) As Variant
Dim YLabel As String, Site As String
'Annual Averages  Charts 3 & 4
Application.ScreenUpdating = False
Sheets("Annual Averages").Select
Sheets("Annual Averages").Range("as5:as57").ClearContents 'clear plot data from previous run         'clear previous results
Sheets("Annual Averages").Range("at5:at57").ClearContents
Sheets("Annual Averages").Range("n3").Select: Parameter1 = ActiveCell.Value
Sheets("Annual Averages").Range("n4").Select: Parameter2 = ActiveCell.Value
For k = 1 To 2
    If k = 1 Then Parameter = Parameter1
    If k = 2 Then Parameter = Parameter2
     If Parameter = "Days>8" Then YLabel = "Number of Days": XMax = 300: XUnit = 50: Range("d48").Select
     If Parameter = "Loss Rate" Then YLabel = "meters/year": XMax = 40: XUnit = 10: Range("f48").Select
     If Parameter = "USGS Flow" Then YLabel = "cfs": XMax = 300: XUnit = 50: Range("z48").Select
     If Parameter = "Rain Inch" Then YLabel = "Inches": XMax = 60: XUnit = 10: Range("aa48").Select
     If Parameter = "BC InFlow" Then YLabel = "mgd": XMax = 15: XUnit = 3: Range("p48").Select
     If Parameter = "PRSFH OutFlow" Then YLabel = "mgd": XMax = 15: XUnit = 3: Range("s48").Select
     If Parameter = "PRSFH Load" Then YLabel = "Pounds": XMax = 600: XUnit = 100: Range("u48").Select
     If Parameter = "BC Load" Then YLabel = "Pounds": XMax = 600: XUnit = 100: Range("r48").Select
     If Parameter = "Lost Fish" Then YLabel = "Pounds": XMax = 300: XUnit = 50: Range("v48").Select
     If Parameter = "Sed Rel" Then YLabel = "Pounds": XMax = 400: XUnit = 100: Range("e48").Select
     If Parameter = "Total Load" Then YLabel = "Pounds": XMax = 12000: XUnit = 3000: Range("w48").Select
     If Parameter = "Lower NP" Then YLabel = "Pounds": XMax = 12000: XUnit = 3000: Range("x48").Select
     If Parameter = "Upper NP" Then YLabel = "Pounds": XMax = 12000: XUnit = 3000: Range("y48").Select
     If Parameter = "Rain Load" Then YLabel = "Pounds": XMax = 600: XUnit = 100: Range("ab48").Select
   If Parameter = "Lake TP" Then YLabel = "TP  mg/m3": XMax = 20: XUnit = 4: Range("c48").Select
   If Parameter = "Stone TP" Then YLabel = "TP  mg/m3": XMax = 32: XUnit = 4: Range("g48").Select
   If Parameter = "Carter TP" Then YLabel = "TP  mg/m3": XMax = 32: XUnit = 4: Range("h48").Select
   If Parameter = "Collision TP" Then YLabel = "TP  mg/m3": XMax = 32: XUnit = 4: Range("i48").Select
   If Parameter = "NB Dead TP" Then YLabel = "TP  mg/m3": XMax = 32: XUnit = 4: Range("j48").Select
   If Parameter = "Vet's TP" Then YLabel = "TP  mg/m3": XMax = 32: XUnit = 4: Range("k48").Select
   If Parameter = "Pioneer TP" Then YLabel = "TP  mg/m3": XMax = 32: XUnit = 4: Range("L48").Select
   If Parameter = "USGS TP" Then YLabel = "TP  mg/m3": XMax = 32: XUnit = 4: Range("m48").Select
   If Parameter = "Haze TP" Then YLabel = "TP  mg/m3": XMax = 32: XUnit = 4: Range("n48").Select
   If Parameter = "NB Ind Hill TP" Then YLabel = "TP  mg/m3": XMax = 32: XUnit = 4: Range("o48").Select
   If Parameter = "PRSFH TP" Then YLabel = "TP  mg/m3": XMax = 32: XUnit = 4: Range("t48").Select
   If Parameter = "BC TP" Then YLabel = "TP  mg/m3": XMax = 20: XUnit = 4: Range("q48").Select
   For j = 1 To 40     ' read'''''''''''''''''''''''''''''''''''''''''''''''''
      TestValue(j) = ActiveCell.Value
      If TestValue(j) = 0 Or TestValue(j) = "" Then Exit For
      ActiveCell.Offset(1, 0).Select
   Next j
   If k = 1 Then Range("as7").Select
   If k = 2 Then Range("at7").Select
   For i = 1 To 40     'print''''''''''''''''''''''''''''''''''''''''''''''''
      ActiveCell.Value = TestValue(i)
      ActiveCell.Offset(1, 0).Select
   Next i
   If k = 1 Then Sheets("Annual Averages").Range("as5").Select: ActiveCell.Value = YLabel:  ActiveCell.Offset(1, 0).Select: ActiveCell.Value = "Annual Average  " & Parameter
   If k = 2 Then Sheets("Annual Averages").Range("at5").Select: ActiveCell.Value = YLabel:  ActiveCell.Offset(1, 0).Select: ActiveCell.Value = "Annual Average  " & Parameter
If k = 1 Then
ActiveSheet.ChartObjects("Chart 12").Activate
        With ActiveChart.Axes(xlValue)
            .MinimumScale = 0
            .MaximumScale = XMax
            .MajorUnit = XUnit
        End With
End If
If k = 2 Then
ActiveSheet.ChartObjects("Chart 11").Activate
        With ActiveChart.Axes(xlValue)
            .MinimumScale = 0
            .MaximumScale = XMax
            .MajorUnit = XUnit
        End With
End If
Next k
Sheets("Annual Averages").Range("n3").Select
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
Worksheets("Annual Averages").Activate
    Sheets("Annual Averages").Select
    TextBox1.Visible = False
    CommandButton4.Caption = "Open"
End Sub