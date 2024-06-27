Attribute VB_Name = "Sheet9"

Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"

Attribute VB_GlobalNameSpace = False

Attribute VB_Creatable = False

Attribute VB_PredeclaredId = True

Attribute VB_Exposed = True

Attribute VB_TemplateDerived = False

Attribute VB_Customizable = True

Attribute VB_Control = "TextBox1, 10, 0, MSForms, TextBox"

Attribute VB_Control = "CommandButton4, 9, 1, MSForms, CommandButton"

Attribute VB_Control = "CommandButton3, 8, 2, MSForms, CommandButton"

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

Worksheets("Trib Flow Corr").Activate

    Sheets("Trib Flow Corr").Select

    TextBox1.Visible = False

    CommandButton4.Caption = "Open"

End Sub

