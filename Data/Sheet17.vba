Attribute VB_Name = "Sheet17"

Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"

Attribute VB_GlobalNameSpace = False

Attribute VB_Creatable = False

Attribute VB_PredeclaredId = True

Attribute VB_Exposed = True

Attribute VB_TemplateDerived = False

Attribute VB_Customizable = True

Attribute VB_Control = "CommandButton1, 2, 0, MSForms, CommandButton"

Attribute VB_Control = "TextBox1, 3, 1, MSForms, TextBox"

Attribute VB_Control = "CommandButton2, 4, 2, MSForms, CommandButton"

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

