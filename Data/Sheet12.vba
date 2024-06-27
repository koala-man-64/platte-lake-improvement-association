Attribute VB_Name = "Sheet12"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Attribute VB_Control = "CommandButton17, 32, 0, MSForms, CommandButton"
Attribute VB_Control = "CommandButton16, 31, 1, MSForms, CommandButton"
Attribute VB_Control = "CommandButton14, 30, 2, MSForms, CommandButton"
Attribute VB_Control = "CommandButton15, 25, 3, MSForms, CommandButton"
Attribute VB_Control = "CommandButton9, 12, 4, MSForms, CommandButton"
Attribute VB_Control = "TextBox1, 9, 5, MSForms, TextBox"
Attribute VB_Control = "CommandButton8, 8, 6, MSForms, CommandButton"
Attribute VB_Control = "CommandButton6, 6, 7, MSForms, CommandButton"
Attribute VB_Control = "CommandButton5, 5, 8, MSForms, CommandButton"
Attribute VB_Control = "CommandButton4, 4, 9, MSForms, CommandButton"
Attribute VB_Control = "CommandButton3, 3, 10, MSForms, CommandButton"
Attribute VB_Control = "CommandButton2, 2, 11, MSForms, CommandButton"
Attribute VB_Control = "CommandButton1, 1, 12, MSForms, CommandButton"
Attribute VB_Control = "CommandButton10, 14, 13, MSForms, CommandButton"
Attribute VB_Control = "CommandButton11, 16, 14, MSForms, CommandButton"
Attribute VB_Control = "CommandButton7, 18, 15, MSForms, CommandButton"
Attribute VB_Control = "CommandButton12, 21, 16, MSForms, CommandButton"
Attribute VB_Control = "CommandButton13, 22, 17, MSForms, CommandButton"
Attribute VB_Control = "CommandButton18, 33, 18, MSForms, CommandButton"
Private Sub CommandButton1_Click()              'lake chem
    Sheets("Lake Chemistry").Select
    TextBox1.Visible = False
    Sheets("Lake Chemistry").Range("h3").Select
End Sub
Private Sub CommandButton18_Click()             'current conditions
Dim Value(20) As Double
Dim LastDate(20) As Date
Dim AlarmLow(20) As Double, AlarmHigh(20)
Dim SelectColumn As Integer
Dim Switch1  As Boolean, Switch2 As Boolean
Application.ScreenUpdating = False
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''  lake chemistry read values
        Sheets("Lake Chemistry").Select
        Sheets("Lake Chemistry").Range("f37").Select: DataCount = ActiveCell.Value
        Sheets("Lake Chemistry").Range("b" + CStr(DataCount + 38)).Select
        LastDate(1) = ActiveCell.Value
        ActiveCell.Offset(0, 4).Select
        Value(1) = ActiveCell.Value
        Sheets("Lake Chemistry").Range("o37").Select: DataCount = ActiveCell.Value
        Sheets("Lake Chemistry").Range("o" + CStr(DataCount + 38)).Select
        Value(2) = ActiveCell.Value
        ActiveCell.Offset(0, -2).Select
        LastDate(2) = ActiveCell.Value
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' lake probe read values
        Switch1 = False
        Switch2 = False
        Sheets("Lake Probe Data").Select
        Sheets("Lake Probe Data").Range("c37").Select: DataCount = ActiveCell.Value
        Sheets("Lake Probe Data").Range("c" + CStr(DataCount + 38)).Select
        For k = 1 To 24                                 'move up up to 24 rows look for depth 0
            Depth = ActiveCell.Value
            If Depth = 90 And Switch1 = False Then
                ActiveCell.Offset(0, -1).Select
                LastDate(4) = ActiveCell.Value
                LastDate(5) = LastDate(4)
                ActiveCell.Offset(0, 2).Select
                Value(4) = ActiveCell.Value
                ActiveCell.Offset(0, 1).Select
                Value(5) = ActiveCell.Value
                ActiveCell.Offset(0, -2).Select
                Switch1 = True
            End If
            If Depth = 0 And Switch2 = False Then
                ActiveCell.Offset(0, -1).Select
                LastDate(3) = ActiveCell.Value
                ActiveCell.Offset(0, 2).Select
                Value(3) = ActiveCell.Value
                ActiveCell.Offset(0, -1).Select
                Switch2 = True
            End If
            If Switch1 = True And Switch2 = True Then Exit For
        ActiveCell.Offset(-1, 0).Select
        Next k
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' stream chemistry read values
        Sheets("Stream Chemistry").Select   'stream data
        For k = 6 To 13     'check columns 3, 6, 9 etc
            SelectColumn = 3 * k - 15       '25 is column Y for 2010
            ActiveSheet.Cells(38, SelectColumn).Select: DataCount = ActiveCell.Value
            ActiveSheet.Cells(DataCount + 39, SelectColumn).Select
            Value(k) = ActiveCell.Value
            ActiveCell.Offset(0, -1).Select
            LastDate(k) = ActiveCell.Value
        Next k
'''''''''''''''''''''''''''''''''''''''''''''''''''''print values
Sheets("Main Menu").Select
Range("w9").Select
For i = 1 To 13
            ActiveCell.Value = Value(i)
            ActiveCell.Offset(0, 1).Select
            ActiveCell.Value = LastDate(i)
            ActiveCell.Offset(1, -1).Select
        Next i
'''''''''''''''''''''''''''''''''''''''''''''''''''''read low & high alarm
Range("y9").Select
   For j = 1 To 13
       AlarmLow(j) = ActiveCell.Value
       ActiveCell.Offset(0, 1).Select
       AlarmHigh(j) = ActiveCell.Value
       ActiveCell.Offset(1, -1).Select
   Next j
 'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
 Range("w9").Select                     'find 3 colors
 For j = 1 To 13
   If j <> 2 And j <> 5 Then
        If Value(j) < AlarmLow(j) Then
            With Selection.Interior
                .Color = 5287936            'green
                .TintAndShade = 0.2
            End With
        ElseIf Value(j) >= AlarmLow(j) And Value(j) < AlarmHigh(j) Then
            With Selection.Interior
                .Color = 65535                         'yellow
                .TintAndShade = 0
            End With
        Else
            With Selection.Interior
                .Color = 255                       'red
                .TintAndShade = 0.4
            End With
        End If
    End If
    If j = 2 Or j = 5 Then
        If Value(j) > AlarmLow(j) Then
            With Selection.Interior
                .Color = 5287936            'green
                .TintAndShade = 0.2
            End With
        ElseIf Value(j) <= AlarmLow(j) And Value(j) > AlarmHigh(j) Then
            With Selection.Interior
                .Color = 65535                         'yellow
                .TintAndShade = 0
            End With
        Else
            With Selection.Interior
                .Color = 255                       'red
                .TintAndShade = 0.4
            End With
        End If
     End If
ActiveCell.Offset(1, 0).Select
Next j
Sheets("Main Menu").Range("c3").Select
End Sub
Private Sub CommandButton2_Click()              'lake probe
    Sheets("Lake Probe Data").Select
    TextBox1.Visible = False
    Sheets("Lake Probe Data").Range("h3").Select
End Sub
Private Sub CommandButton17_Click()             'support
    Sheets("Support").Select
    TextBox1.Visible = False
    Sheets("Support").Range("h3").Select
End Sub
Private Sub CommandButton3_Click()              'stream chem
    Sheets("Stream Chemistry").Select
    TextBox1.Visible = False
    Sheets("Stream Chemistry").Range("i4").Select
End Sub
Private Sub CommandButton12_Click()             'stream probe
Sheets("Stream Probe").Select
    TextBox1.Visible = False
    Sheets("Stream Probe").Range("j3").Select
End Sub
Private Sub CommandButton5_Click()              'near shore
    Sheets("Near-Shore").Select
    TextBox1.Visible = False
    Sheets("Near-Shore").Range("i6").Select
End Sub
Private Sub CommandButton9_Click()              'wet TP
    Sheets("Wet Weather TP").Select
    TextBox1.Visible = False
    Sheets("Wet Weather TP").Range("i4").Select
End Sub
Private Sub CommandButton15_Click()             'rain flow comparision
    Sheets("Flow & Rain & TP Comparison").Select
    TextBox1.Visible = False
    Sheets("Flow & Rain & TP Comparison").Range("i4").Select
End Sub
Private Sub CommandButton14_Click()             'rain flow data
Sheets("Flow & Rain Data").Select
    TextBox1.Visible = False
    Sheets("Flow & Rain Data").Range("k3").Select
End Sub
Private Sub CommandButton7_Click()              'trib flow corr
    Sheets("Trib Flow Corr").Select
    TextBox1.Visible = False
    Sheets("Trib Flow Corr").Range("g3").Select
End Sub
Private Sub CommandButton10_Click()             'moving avg
Sheets("Moving Average").Select
    TextBox1.Visible = False
    Sheets("Moving Average").Range("L5").Select
End Sub
Private Sub CommandButton4_Click()              'long term trend
    Sheets("Long-Term Trends").Select
    TextBox1.Visible = False
    Sheets("Long-Term Trends").Range("h3").Select
End Sub
Private Sub CommandButton16_Click()             'ann avg
    Sheets("Annual Averages").Select
    TextBox1.Visible = False
    Sheets("Annual Averages").Range("g3").Select
End Sub
Private Sub CommandButton11_Click()             'mass bal
Sheets("Watershed Mass Bal").Select
    TextBox1.Visible = False
    Sheets("Watershed Mass Bal").Range("g3").Select
End Sub
Private Sub CommandButton13_Click()             'lake model
Sheets("Lake TP Model").Select
    TextBox1.Visible = False
    Sheets("Lake TP Model").Range("e12").Select
End Sub
Private Sub CommandButton6_Click()              'miscellaneous
    Sheets("Miscellaneous").Select
    TextBox1.Visible = False
    Sheets("Miscellaneous").Range("a1").Select
End Sub
Private Sub commandbutton8_click()              'documentation
    If CommandButton8.Caption = "Open" Then
        CommandButton8.Caption = "Close"
        TextBox1.Visible = True
    Else
        CommandButton8.Caption = "Open"
        TextBox1.Visible = False
    End If
End Sub