Attribute VB_Name = "Sheet14"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Attribute VB_Control = "TextBox1, 5, 0, MSForms, TextBox"
Attribute VB_Control = "CommandButton1, 1, 1, MSForms, CommandButton"
Attribute VB_Control = "CommandButton2, 2, 2, MSForms, CommandButton"
Private Sub CommandButton1_Click()
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
Private Sub worksheet_Activate()
Worksheets("Lake TP Model").Activate
    Sheets("Lake TP Model").Select
    TextBox1.Visible = False
    CommandButton2.Caption = "Open"
End Sub
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
If ActiveCell.Address = "$X$14" Then
        Range("x14:x17").Interior.ColorIndex = 2
        ActiveCell.Interior.ColorIndex = 15
        ActiveSheet.ChartObjects("Chart 11").Visible = True
        ActiveSheet.ChartObjects("Chart 6").Visible = False     'was 6
        ActiveSheet.ChartObjects("Chart 10").Visible = False
        ActiveSheet.ChartObjects("Chart 1").Visible = False     'was 5
End If
If ActiveCell.Address = "$X$15" Then
        Range("x14:x17").Interior.ColorIndex = 2
        ActiveCell.Interior.ColorIndex = 15
        ActiveSheet.ChartObjects("Chart 11").Visible = False
        ActiveSheet.ChartObjects("Chart 6").Visible = True
        ActiveSheet.ChartObjects("Chart 10").Visible = False
        ActiveSheet.ChartObjects("Chart 1").Visible = False
End If
If ActiveCell.Address = "$X$16" Then
        Range("x14:x17").Interior.ColorIndex = 2
        ActiveCell.Interior.ColorIndex = 15
        ActiveSheet.ChartObjects("Chart 11").Visible = False
        ActiveSheet.ChartObjects("Chart 6").Visible = False
        ActiveSheet.ChartObjects("Chart 10").Visible = True
        ActiveSheet.ChartObjects("Chart 1").Visible = False
End If
If ActiveCell.Address = "$X$17" Then
        Range("x14:x17").Interior.ColorIndex = 2
        ActiveCell.Interior.ColorIndex = 15
        ActiveSheet.ChartObjects("Chart 11").Visible = False
        ActiveSheet.ChartObjects("Chart 6").Visible = False
        ActiveSheet.ChartObjects("Chart 10").Visible = False
        ActiveSheet.ChartObjects("Chart 1").Visible = True
End If
End Sub