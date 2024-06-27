Attribute VB_Name = "Sheet5"

Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"

Attribute VB_GlobalNameSpace = False

Attribute VB_Creatable = False

Attribute VB_PredeclaredId = True

Attribute VB_Exposed = True

Attribute VB_TemplateDerived = False

Attribute VB_Customizable = True

Attribute VB_Control = "CommandButton2, 35, 0, MSForms, CommandButton"

Attribute VB_Control = "TextBox1, 29, 1, MSForms, TextBox"

Attribute VB_Control = "CommandButton1, 17, 2, MSForms, CommandButton"

Attribute VB_Control = "CommandButton3, 32, 3, MSForms, CommandButton"

Attribute VB_Control = "TextBox2, 38, 4, MSForms, TextBox"

Attribute VB_Control = "CommandButton4, 39, 5, MSForms, CommandButton"

Private Sub CommandButton1_Click()

Sheets("Main Menu").Select

Sheets("Main Menu").Range("g11").Select



End Sub

Private Sub CommandButton2_Click()

If CommandButton2.Caption = "Open" Then 'documentation

        CommandButton2.Caption = "Close"

        TextBox2.Visible = True

    Else

        CommandButton2.Caption = "Open"

        TextBox2.Visible = False

    End If

End Sub

Private Sub CommandButton3_Click()

If CommandButton3.Caption = "Open" Then 'references

        CommandButton3.Caption = "Close"

        TextBox1.Visible = True

    Else

        CommandButton3.Caption = "Open"

        TextBox1.Visible = False

    End If

End Sub

Private Sub CommandButton4_Click()

Dim Value1(50) As Integer

Dim Value2(50) As Double, Value3(50) As Double

Sheets("Miscellaneous").Select

Sheets("Miscellaneous").Range("ab66").Select

flowcount = ActiveCell.Value



Application.ScreenUpdating = False

Sheets("Miscellaneous").Range("at21:av59").ClearContents



Sheets("Miscellaneous").Range("aa30").Select

For i = 1 To flowcount

  Value1(i) = ActiveCell.Value

  ActiveCell.Offset(0, 5).Select

  Value2(i) = ActiveCell.Value

  ActiveCell.Offset(0, 1).Select

  Value3(i) = ActiveCell.Value

  

  ActiveCell.Offset(1, -6).Select

  

Next i



Sheets("Miscellaneous").Range("at21").Select

For i = 1 To flowcount

   ActiveCell.Value = Value1(i)

  ActiveCell.Offset(0, 1).Select

   ActiveCell.Value = Value2(i)

  ActiveCell.Offset(0, 1).Select

   ActiveCell.Value = Value3(i)

  ActiveCell.Offset(1, -2).Select

Next i

Application.Goto Range("A1"), True

ActiveWindow.VisibleRange(1, 1).Select

Sheets("Miscellaneous").Range("ah27").Select



End Sub



Private Sub worksheet_Activate()

Worksheets("Miscellaneous").Activate

    Sheets("Miscellaneous").Select

    TextBox2.Visible = False

    CommandButton2.Caption = "Open"

    TextBox1.Visible = False

    CommandButton3.Caption = "Open"

End Sub



