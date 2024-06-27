Attribute VB_Name = "Sheet13"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Attribute VB_Control = "CommandButton3, 10, 0, MSForms, CommandButton"
Attribute VB_Control = "CommandButton2, 8, 1, MSForms, CommandButton"
Attribute VB_Control = "HTMLOption3, 1, 2, MSForms, HTMLOption"
Attribute VB_Control = "HTMLOption2, 2, 3, MSForms, HTMLOption"
Attribute VB_Control = "HTMLOption1, 3, 4, MSForms, HTMLOption"
Attribute VB_Control = "CommandButton1, 5, 5, MSForms, CommandButton"
Attribute VB_Control = "CommandButton4, 13, 6, MSForms, CommandButton"
Attribute VB_Control = "TextBox1, 14, 7, MSForms, TextBox"
Private Sub CommandButton1_Click()
Dim i As Integer, SelectYear As Integer, FebDays As Integer, DayCount As Integer
Dim Rain(2000, 2000) As Double, AnnualTotalRain(3000) As Variant
Dim Row As String
Dim k As Integer, MaxYear As Integer
Application.ScreenUpdating = False
Sheets("Flow & Rain Data").Range("k3").Select
SelectYear = ActiveCell.Value
Sheets("Flow & Rain Data").Range("bq3").Select
FebDays = ActiveCell.Value
Sheets("Flow & Rain Data").Range("br3").Select
LastRainYear = ActiveCell.Value
If LastRainYear >= SelectYear Then
   response% = MsgBox("Rain Data for " & CStr(SelectYear) & " have already been entered.", 64)
   Exit Sub
End If
If LastRainYear <= SelectYear Then
     answer = MsgBox("Have new Rain data for " & CStr(SelectYear) & " been pasted into cells B13, B49, and B85?", vbQuestion + vbYesNo)
     If answer = 7 Then Exit Sub    'no
     If answer = 6 Then continue = True       'yes
End If
For k = 1 To 3
    If k = 1 Then Row = "13"
    If k = 2 Then Row = "49"
    If k = 3 Then Row = "85"
    DayCount = 0
    Sheets("Flow & Rain Data").Range("c" + Row).Select
    For i = 1 To 31 'jan
        DayCount = DayCount + 1
        If IsNumeric(ActiveCell.Value) Then
                Rain(k, DayCount) = ActiveCell.Value
            Else
                Rain(k, DayCount) = 0
        End If
        ActiveCell.Offset(1, 0).Select
    Next i
    Sheets("Flow & Rain Data").Range("d" + Row).Select
    For i = 1 To FebDays 'feb
        DayCount = DayCount + 1
        If IsNumeric(ActiveCell.Value) Then
                Rain(k, DayCount) = ActiveCell.Value
            Else
                Rain(k, DayCount) = 0
        End If
        ActiveCell.Offset(1, 0).Select
    Next i
    Sheets("Flow & Rain Data").Range("e" + Row).Select
    For i = 1 To 31 'mar
        DayCount = DayCount + 1
        If IsNumeric(ActiveCell.Value) Then
                Rain(k, DayCount) = ActiveCell.Value
            Else
                Rain(k, DayCount) = 0
        End If
        ActiveCell.Offset(1, 0).Select
    Next i
    Sheets("Flow & Rain Data").Range("f" + Row).Select
    For i = 1 To 30 'apr
        DayCount = DayCount + 1
        If IsNumeric(ActiveCell.Value) Then
                Rain(k, DayCount) = ActiveCell.Value
            Else
                Rain(k, DayCount) = 0
        End If
        ActiveCell.Offset(1, 0).Select
    Next i
    Sheets("Flow & Rain Data").Range("g" + Row).Select
    For i = 1 To 31 'may
        DayCount = DayCount + 1
        If IsNumeric(ActiveCell.Value) Then
                Rain(k, DayCount) = ActiveCell.Value
            Else
                Rain(k, DayCount) = 0
        End If
        ActiveCell.Offset(1, 0).Select
    Next i
    Sheets("Flow & Rain Data").Range("h" + Row).Select
    For i = 1 To 30 'jun
        DayCount = DayCount + 1
        If IsNumeric(ActiveCell.Value) Then
                Rain(k, DayCount) = ActiveCell.Value
            Else
                Rain(k, DayCount) = 0
        End If
        ActiveCell.Offset(1, 0).Select
    Next i
    Sheets("Flow & Rain Data").Range("i" + Row).Select
    For i = 1 To 31 'jly
        DayCount = DayCount + 1
        If IsNumeric(ActiveCell.Value) Then
                Rain(k, DayCount) = ActiveCell.Value
            Else
                Rain(k, DayCount) = 0
        End If
        ActiveCell.Offset(1, 0).Select
    Next i
    Sheets("Flow & Rain Data").Range("j" + Row).Select
    For i = 1 To 31 'aug
        DayCount = DayCount + 1
        If IsNumeric(ActiveCell.Value) Then
                Rain(k, DayCount) = ActiveCell.Value
            Else
                Rain(k, DayCount) = 0
        End If
        ActiveCell.Offset(1, 0).Select
    Next i
    Sheets("Flow & Rain Data").Range("k" + Row).Select
    For i = 1 To 30 'sep
        DayCount = DayCount + 1
        If IsNumeric(ActiveCell.Value) Then
                Rain(k, DayCount) = ActiveCell.Value
            Else
                Rain(k, DayCount) = 0
        End If
        ActiveCell.Offset(1, 0).Select
    Next i
    Sheets("Flow & Rain Data").Range("l" + Row).Select
    For i = 1 To 31 'oct
        DayCount = DayCount + 1
        If IsNumeric(ActiveCell.Value) Then
                Rain(k, DayCount) = ActiveCell.Value
            Else
                Rain(k, DayCount) = 0
        End If
        ActiveCell.Offset(1, 0).Select
    Next i
    Sheets("Flow & Rain Data").Range("m" + Row).Select
    For i = 1 To 30 'nov
        DayCount = DayCount + 1
        If IsNumeric(ActiveCell.Value) Then
                Rain(k, DayCount) = ActiveCell.Value
            Else
                Rain(k, DayCount) = 0
        End If
        ActiveCell.Offset(1, 0).Select
    Next i
    Sheets("Flow & Rain Data").Range("n" + Row).Select
    For i = 1 To 31 'dec
        DayCount = DayCount + 1
        If IsNumeric(ActiveCell.Value) Then
                Rain(k, DayCount) = ActiveCell.Value
            Else
                Rain(k, DayCount) = 0
        End If
        ActiveCell.Offset(1, 0).Select
    Next i
Next k
PrintRow = 25 + (SelectYear - 2010)         '25 is column Y for 2010
ActiveSheet.Cells(12, PrintRow).Select
ActiveCell.Value = SelectYear
ActiveSheet.Cells(13, PrintRow).Select
For j = 1 To DayCount
    ActiveCell.Value = (Rain(1, j) + Rain(2, j) + Rain(3, j)) / 3
    ActiveCell.Offset(1, 0).Select
Next j
Sheets("Flow & Rain Data").Range("y11").Select      'read annual total
For k = 2010 To 2050
    AnnualTotalRain(k) = ActiveCell.Value
    ActiveCell.Offset(0, 1).Select
Next k
Sheets("Annual Averages").Select                    'print annual total
Sheets("Annual Averages").Range("aa48").Select
For k = 2010 To 2050
    ActiveCell.Value = AnnualTotalRain(k)
    ActiveCell.Offset(1, 0).Select
Next k
response% = MsgBox("New Rain Data for " & CStr(SelectYear) & " have been entered.", 64)
Sheets("Flow & Rain Data").Select
Sheets("Flow & Rain Data").Range("k3").Select
End Sub
Private Sub CommandButton4_Click()
Dim Flow(370) As Double, AvgAnnualFlow(3000) As Variant
Application.ScreenUpdating = False
Sheets("Flow & Rain Data").Select
Sheets("Flow & Rain Data").Range("q3").Select
SelectYear = ActiveCell.Value
Sheets("Flow & Rain Data").Range("bt3").Select
FebDays = ActiveCell.Value
Sheets("Flow & Rain Data").Range("bs3").Select
LastFlowYear = ActiveCell.Value
If LastFlowYear >= SelectYear Then
   response% = MsgBox("Flow Data for " & CStr(SelectYear) & " have already been entered.", 64)
   Exit Sub
End If
If LastFlowYear <= SelectYear Then
     answer = MsgBox("Have old Flow data been cleared and new Flow data for " & CStr(SelectYear) & " been pasted into cell Q13 and converted from Text to Columns?", vbQuestion + vbYesNo)
     If answer = 7 Then Exit Sub    'no
     If answer = 6 Then continue = True   'yes
End If
Sheets("Flow & Rain Data").Range("t13").Select
For i = 1 To 337 + FebDays
    Flow(i) = ActiveCell.Value
    ActiveCell.Offset(1, 0).Select
Next i
PrintRow = 25 + (SelectYear - 1990)
ActiveSheet.Cells(382, PrintRow).Select
ActiveCell.Value = SelectYear
ActiveSheet.Cells(383, PrintRow).Select
For j = 1 To 337 + FebDays
    ActiveCell.Value = Flow(j)
    ActiveCell.Offset(1, 0).Select
Next j
Sheets("Flow & Rain Data").Range("as381").Select      'read annual avg flow
For k = 2010 To 2050
    AvgAnnualFlow(k) = ActiveCell.Value
    ActiveCell.Offset(0, 1).Select
Next k
Sheets("Annual Averages").Select
Sheets("Annual Averages").Range("z48").Select
For k = 2010 To 2050
    ActiveCell.Value = AvgAnnualFlow(k)
    ActiveCell.Offset(1, 0).Select
Next k
response% = MsgBox("New Flow Data for " & CStr(SelectYear) & " have been entered.", 64)
Sheets("Flow & Rain Data").Select
Sheets("Flow & Rain Data").Range("k3").Select
End Sub
Private Sub worksheet_Activate()
Worksheets("Flow & Rain Data").Activate
    Sheets("Flow & Rain Data").Select
    TextBox1.Visible = False
    CommandButton2.Caption = "Open"
End Sub
Private Sub CommandButton3_Click()
    Sheets("Main Menu").Select
    Sheets("Main Menu").Range("g11").Select
End Sub
Private Sub CommandButton2_Click()
If CommandButton2.Caption = "Open" Then
        CommandButton2.Caption = "Close"
        TextBox1.Visible = True
    Else
        CommandButton2.Caption = "Open"
        TextBox1.Visible = False
    End If
End Sub