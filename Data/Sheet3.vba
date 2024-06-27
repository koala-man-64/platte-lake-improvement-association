Attribute VB_Name = "Sheet3"

Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"

Attribute VB_GlobalNameSpace = False

Attribute VB_Creatable = False

Attribute VB_PredeclaredId = True

Attribute VB_Exposed = True

Attribute VB_TemplateDerived = False

Attribute VB_Customizable = True

Attribute VB_Control = "CommandButton2, 2, 1, MSForms, CommandButton"

Attribute VB_Control = "CommandButton1, 7, 2, MSForms, CommandButton"

Attribute VB_Control = "CommandButton3, 8, 3, MSForms, CommandButton"

Attribute VB_Control = "TextBox1, 10, 4, MSForms, TextBox"

Attribute VB_Control = "CommandButton4, 15, 5, MSForms, CommandButton"

Private Sub CommandButton1_Click()

Sheets("Main Menu").Select

    Sheets("Main Menu").Range("g11").Select

End Sub

Private Sub CommandButton2_Click()          ''''''''''''''''''''''''''''''chart 1



Dim M1(500) As Double, M2(500) As Double, M3(500) As Double

Dim M4(500) As Double, M5(500) As Double, M6(500) As Double

Dim M7(500) As Double, M8(500) As Double, M9(500) As Double

Dim M10(500) As Double, M11(500) As Double, M12(500) As Double



Dim Count1 As Integer, Count2 As Integer, Count3 As Integer

Dim Count4 As Integer, Count5 As Integer, Count6 As Integer

Dim Count7 As Integer, Count8 As Integer, Count9 As Integer

Dim Count10 As Integer, Count11 As Integer, Count12 As Integer



Dim TestDate As Date, TestValue As Double

Dim PlotDate1(20000) As Date, PlotValue1(20000) As Double

Dim PlotDate2(20000) As Date, PlotValue2(20000) As Double

Dim PlotDate3(20000) As Date, PlotValue3(20000) As Double

Dim HDelta As Integer

Dim Site As String

Dim DataCount As Integer

Dim CompareDate(500) As Date, CompareValue(500) As Double, CompareYear As Integer



Application.ScreenUpdating = False



Sheets("Moving Average").Select

Sheets("Moving Average").Range("b41:b133").ClearContents 'clear plot data from previous run         'clear previous results

Sheets("Moving Average").Range("d41:d133").ClearContents 'clear plot data from previous run

Sheets("Moving Average").Range("g41:r133").ClearContents 'clear plot data from previous run       'clear previous results



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



Sheets("Moving Average").Range("j2").Select: StartYear = ActiveCell.Value

Sheets("Moving Average").Range("j3").Select: EndYear = ActiveCell.Value         'year

Sheets("Moving Average").Range("j5").Select: Site = ActiveCell.Value            'site

Sheets("Moving Average").Range("j7").Select: CompareYear = ActiveCell.Value    'site



If EndYear < StartYear Then

    response% = MsgBox("The End Year must be greater than or equal to the Start Year.", 64)

    Exit Sub

End If



Count1 = 1: Count2 = 1: Count3 = 1: Count4 = 1 'xxxxxxxxx

Count5 = 1: Count6 = 1: Count7 = 1: Count8 = 1

Count9 = 1: Count10 = 1: Count11 = 1: Count12 = 1

CompareCount = 1



If Site = "Lake TP" Then

    HDelta = 4

    Sheets("Lake Chemistry").Select

    Sheets("Lake Chemistry").Range("f37").Select: DataCount = ActiveCell.Value

    Sheets("Lake Chemistry").Range("b39").Select

End If



If Site = "Secchi" Then

    HDelta = 2

    Sheets("Lake Chemistry").Select

    Sheets("Lake Chemistry").Range("o37").Select: DataCount = ActiveCell.Value

    Sheets("Lake Chemistry").Range("m39").Select

End If



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



If Site = "Stone" Then

    HDelta = 1

    Sheets("Stream Chemistry").Select

    Sheets("Stream Chemistry").Range("c38").Select: DataCount = ActiveCell.Value

    Sheets("Stream Chemistry").Range("b40").Select

End If



If Site = "Vet's" Then

    HDelta = 1

    Sheets("Stream Chemistry").Select

    Sheets("Stream Chemistry").Range("f38").Select: DataCount = ActiveCell.Value

    Sheets("Stream Chemistry").Range("e40").Select

End If



If Site = "Haze" Then

    HDelta = 1

    Sheets("Stream Chemistry").Select

    Sheets("Stream Chemistry").Range("i38").Select: DataCount = ActiveCell.Value

    Sheets("Stream Chemistry").Range("h40").Select

End If



If Site = "Carter" Then

    HDelta = 1

    Sheets("Stream Chemistry").Select

    Sheets("Stream Chemistry").Range("l38").Select: DataCount = ActiveCell.Value

    Sheets("Stream Chemistry").Range("k40").Select

End If



If Site = "Pioneer" Then

    HDelta = 1

    Sheets("Stream Chemistry").Select

    Sheets("Stream Chemistry").Range("o38").Select: DataCount = ActiveCell.Value

    Sheets("Stream Chemistry").Range("n40").Select

End If



If Site = "USGS" Then

    HDelta = 1

    Sheets("Stream Chemistry").Select

    Sheets("Stream Chemistry").Range("r38").Select: DataCount = ActiveCell.Value

    Sheets("Stream Chemistry").Range("q40").Select

End If



If Site = "NB Ind Hill" Then

    HDelta = 1

    Sheets("Stream Chemistry").Select

    Sheets("Stream Chemistry").Range("u38").Select: DataCount = ActiveCell.Value

    Sheets("Stream Chemistry").Range("t40").Select

End If



If Site = "NB Dead" Then

    HDelta = 1

    Sheets("Stream Chemistry").Select

    Sheets("Stream Chemistry").Range("x38").Select: DataCount = ActiveCell.Value

    Sheets("Stream Chemistry").Range("w40").Select

End If



If Site = "NB Hooker" Then

    HDelta = 1

    Sheets("Stream Chemistry").Select

    Sheets("Stream Chemistry").Range("aa38").Select: DataCount = ActiveCell.Value

    Sheets("Stream Chemistry").Range("z40").Select

End If



If Site = "M22" Then

    HDelta = 1

    Sheets("Stream Chemistry").Select

    Sheets("Stream Chemistry").Range("ad38").Select: DataCount = ActiveCell.Value

    Sheets("Stream Chemistry").Range("ac40").Select

End If



If Site = "BC Old Res" Then

    HDelta = 1

    Sheets("Stream Chemistry").Select

    Sheets("Stream Chemistry").Range("ag38").Select: DataCount = ActiveCell.Value

    Sheets("Stream Chemistry").Range("af40").Select

End If



If Site = "Collision" Then

    HDelta = 1

    Sheets("Stream Chemistry").Select

    Sheets("Stream Chemistry").Range("aj38").Select: DataCount = ActiveCell.Value

    Sheets("Stream Chemistry").Range("ai40").Select

End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



For i = 1 To DataCount

    

    If Month(ActiveCell.Value) = 1 And Year(ActiveCell.Value) >= StartYear And Year(ActiveCell.Value) <= EndYear Then

        ActiveCell.Offset(0, HDelta).Select

        M1(Count1) = ActiveCell.Value

        Count1 = Count1 + 1

        ActiveCell.Offset(1, -HDelta).Select

    End If



    If Month(ActiveCell.Value) = 2 And Year(ActiveCell.Value) >= StartYear And Year(ActiveCell.Value) <= EndYear Then

        ActiveCell.Offset(0, HDelta).Select

        M2(Count2) = ActiveCell.Value

        Count2 = Count2 + 1

        ActiveCell.Offset(0, -HDelta).Select

    End If



    If Month(ActiveCell.Value) = 3 And Year(ActiveCell.Value) >= StartYear And Year(ActiveCell.Value) <= EndYear Then

        ActiveCell.Offset(0, HDelta).Select

        M3(Count3) = ActiveCell.Value

        Count3 = Count3 + 1

        ActiveCell.Offset(0, -HDelta).Select

    End If

    

    If Month(ActiveCell.Value) = HDelta And Year(ActiveCell.Value) >= StartYear And Year(ActiveCell.Value) <= EndYear Then

        ActiveCell.Offset(0, HDelta).Select

        M4(Count4) = ActiveCell.Value

        Count4 = Count4 + 1

        ActiveCell.Offset(0, -HDelta).Select

    End If



    If Month(ActiveCell.Value) = 5 And Year(ActiveCell.Value) >= StartYear And Year(ActiveCell.Value) <= EndYear Then

        ActiveCell.Offset(0, HDelta).Select

        M5(Count5) = ActiveCell.Value

        Count5 = Count5 + 1

        ActiveCell.Offset(0, -HDelta).Select

    End If

    

    If Month(ActiveCell.Value) = 6 And Year(ActiveCell.Value) >= StartYear And Year(ActiveCell.Value) <= EndYear Then

        ActiveCell.Offset(0, HDelta).Select

        M6(Count6) = ActiveCell.Value

        Count6 = Count6 + 1

        ActiveCell.Offset(0, -HDelta).Select

    End If



    If Month(ActiveCell.Value) = 7 And Year(ActiveCell.Value) >= StartYear And Year(ActiveCell.Value) <= EndYear Then

        ActiveCell.Offset(0, HDelta).Select

        M7(Count7) = ActiveCell.Value

        Count7 = Count7 + 1

        ActiveCell.Offset(0, -HDelta).Select

    End If



    If Month(ActiveCell.Value) = 8 And Year(ActiveCell.Value) >= StartYear And Year(ActiveCell.Value) <= EndYear Then

        ActiveCell.Offset(0, HDelta).Select

        M8(Count8) = ActiveCell.Value

        Count8 = Count8 + 1

        ActiveCell.Offset(0, -HDelta).Select

    End If



    If Month(ActiveCell.Value) = 9 And Year(ActiveCell.Value) >= StartYear And Year(ActiveCell.Value) <= EndYear Then

        ActiveCell.Offset(0, HDelta).Select

        M9(Count9) = ActiveCell.Value

        Count9 = Count9 + 1

        ActiveCell.Offset(0, -HDelta).Select

    End If



    If Month(ActiveCell.Value) = 10 And Year(ActiveCell.Value) >= StartYear And Year(ActiveCell.Value) <= EndYear Then

        ActiveCell.Offset(0, HDelta).Select

        M10(Count10) = ActiveCell.Value

        Count10 = Count10 + 1

        ActiveCell.Offset(0, -HDelta).Select

    End If



    If Month(ActiveCell.Value) = 11 And Year(ActiveCell.Value) >= StartYear And Year(ActiveCell.Value) <= EndYear Then

        ActiveCell.Offset(0, HDelta).Select

        M11(Count11) = ActiveCell.Value

        Count11 = Count11 + 1

        ActiveCell.Offset(0, -HDelta).Select

    End If

    

    If Month(ActiveCell.Value) = 12 And Year(ActiveCell.Value) >= StartYear And Year(ActiveCell.Value) <= EndYear Then

        ActiveCell.Offset(0, HDelta).Select

        M12(Count12) = ActiveCell.Value

        Count12 = Count12 + 1

        ActiveCell.Offset(0, -HDelta).Select

    End If

 

If Year(ActiveCell.Value) = CompareYear Then

        CompareDate(CompareCount) = ActiveCell.Value

        ActiveCell.Offset(0, HDelta).Select

        CompareValue(CompareCount) = ActiveCell.Value

        CompareCount = CompareCount + 1

        ActiveCell.Offset(0, -HDelta).Select

End If



ActiveCell.Offset(1, 0).Select



Next i





'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx

Sheets("Moving Average").Select

Sheets("Moving Average").Range("g41").Select

For i = 1 To Count1 - 1

    ActiveCell.Value = M1(i)

    ActiveCell.Offset(1, 0).Select

Next i



Sheets("Moving Average").Range("h41").Select

For i = 1 To Count2 - 1

    ActiveCell.Value = M2(i)

    ActiveCell.Offset(1, 0).Select

Next i



Sheets("Moving Average").Range("i41").Select

For i = 1 To Count3 - 1

    ActiveCell.Value = M3(i)

    ActiveCell.Offset(1, 0).Select

Next i



Sheets("Moving Average").Range("j41").Select

For i = 1 To Count4 - 1

    ActiveCell.Value = M4(i)

    ActiveCell.Offset(1, 0).Select

Next i



Sheets("Moving Average").Range("k41").Select

For i = 1 To Count5 - 1

    ActiveCell.Value = M5(i)

    ActiveCell.Offset(1, 0).Select

Next i



Sheets("Moving Average").Range("l41").Select

For i = 1 To Count6 - 1

    ActiveCell.Value = M6(i)

    ActiveCell.Offset(1, 0).Select

Next i



Sheets("Moving Average").Range("m41").Select

For i = 1 To Count7 - 1

    ActiveCell.Value = M7(i)

    ActiveCell.Offset(1, 0).Select

Next i



Sheets("Moving Average").Range("n41").Select

For i = 1 To Count8 - 1

    ActiveCell.Value = M8(i)

    ActiveCell.Offset(1, 0).Select

Next i



Sheets("Moving Average").Range("o41").Select

For i = 1 To Count9 - 1

    ActiveCell.Value = M9(i)

    ActiveCell.Offset(1, 0).Select

Next i



Sheets("Moving Average").Range("p41").Select

For i = 1 To Count10 - 1

    ActiveCell.Value = M10(i)

    ActiveCell.Offset(1, 0).Select

Next i



Sheets("Moving Average").Range("q41").Select

For i = 1 To Count11 - 1

    ActiveCell.Value = M11(i)

    ActiveCell.Offset(1, 0).Select

Next i



Sheets("Moving Average").Range("r41").Select

For i = 1 To Count12 - 1

    ActiveCell.Value = M12(i)

    ActiveCell.Offset(1, 0).Select

Next i



Sheets("Moving Average").Range("b41").Select:

For i = 1 To CompareCount - 1

        ActiveCell.Value = CompareDate(i)

        ActiveCell.Offset(0, 2).Select

        ActiveCell.Value = CompareValue(i)

        CompareCount = CompareCount + 1

        ActiveCell.Offset(0, -2).Select

ActiveCell.Offset(1, 0).Select



Next i



'Sheets("Moving Average").Range("d39").Select  'find maxiumum

 Sheets("Moving Average").Range("b10").Select  'find maxiumum

     MaxValue = ActiveCell.Value

     

     ActiveSheet.ChartObjects("Chart 1").Activate

     ActiveChart.ChartArea.Select

  

     If MaxValue = 16 Then

        With ActiveChart.Axes(xlValue)

            .MinimumScale = 0

            .MaximumScale = MaxValue

            .MajorUnit = 4

        End With

     End If

       

     If MaxValue = 24 Then

        With ActiveChart.Axes(xlValue)

            .MinimumScale = 0

            .MaximumScale = MaxValue

            .MajorUnit = 4

        End With

     End If

     

     If MaxValue = 48 Then

        With ActiveChart.Axes(xlValue)

            .MinimumScale = 0

            .MaximumScale = MaxValue

            .MajorUnit = 8

        End With

     End If

     

     If MaxValue = 96 Then

        With ActiveChart.Axes(xlValue)

            .MinimumScale = 0

            .MaximumScale = MaxValue

            .MajorUnit = 16

        End With

     End If





If Site = "Secchi" Then Sheets("Moving Average").Range("ap9").Select: ActiveCell.Value = "Secchi  Feet"        'chart title

If Site <> "Secchi" Then Sheets("Moving Average").Range("ap9").Select: ActiveCell.Value = "Toral P    mg/m3"        'chart title





Sheets("Moving Average").Range("ap2").Select: ActiveCell.Value = Site & "  " + CStr(StartYear) & " to " + CStr(EndYear)    'define chart title

Sheets("Moving Average").Range("ap3").Select: ActiveCell.Value = "Compared to " + CStr(CompareYear)    'define chart title



Sheets("Moving Average").Range("j3").Select



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



Private Sub CommandButton4_Click()                  '''''''''''''''''''''''''''''''chart 2





Dim M1(500) As Double, M2(500) As Double, M3(500) As Double

Dim M4(500) As Double, M5(500) As Double, M6(500) As Double

Dim M7(500) As Double, M8(500) As Double, M9(500) As Double

Dim M10(500) As Double, M11(500) As Double, M12(500) As Double



Dim Count1 As Integer, Count2 As Integer, Count3 As Integer

Dim Count4 As Integer, Count5 As Integer, Count6 As Integer

Dim Count7 As Integer, Count8 As Integer, Count9 As Integer

Dim Count10 As Integer, Count11 As Integer, Count12 As Integer



Dim TestDate As Date, TestValue As Double

Dim PlotDate1(20000) As Date, PlotValue1(20000) As Double

Dim PlotDate2(20000) As Date, PlotValue2(20000) As Double

Dim PlotDate3(20000) As Date, PlotValue3(20000) As Double

Dim HDelta As Integer

Dim Site As String

Dim DataCount As Integer

Dim CompareDate(500) As Date, CompareValue(500) As Double, CompareYear As Integer



Application.ScreenUpdating = False



Sheets("Moving Average").Select

Sheets("Moving Average").Range("t41:t133").ClearContents 'clear plot data from previous run         'clear previous results

Sheets("Moving Average").Range("v41:v133").ClearContents 'clear plot data from previous run

Sheets("Moving Average").Range("y41:aj133").ClearContents 'clear plot data from previous run       'clear previous results



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



Sheets("Moving Average").Range("n2").Select: StartYear = ActiveCell.Value

Sheets("Moving Average").Range("n3").Select: EndYear = ActiveCell.Value         'year

Sheets("Moving Average").Range("n5").Select: Site = ActiveCell.Value            'site

Sheets("Moving Average").Range("n7").Select: CompareYear = ActiveCell.Value    'site



If EndYear < StartYear Then

    response% = MsgBox("The End Year must be greater than or equal to the Start Year.", 64)

    Exit Sub

End If



Count1 = 1: Count2 = 1: Count3 = 1: Count4 = 1 'xxxxxxxxx

Count5 = 1: Count6 = 1: Count7 = 1: Count8 = 1

Count9 = 1: Count10 = 1: Count11 = 1: Count12 = 1

CompareCount = 1



If Site = "Lake TP" Then

    HDelta = 4

    Sheets("Lake Chemistry").Select

    Sheets("Lake Chemistry").Range("f37").Select: DataCount = ActiveCell.Value

    Sheets("Lake Chemistry").Range("b39").Select

End If



If Site = "Secchi" Then

    HDelta = 2

    Sheets("Lake Chemistry").Select

    Sheets("Lake Chemistry").Range("o37").Select: DataCount = ActiveCell.Value

    Sheets("Lake Chemistry").Range("m39").Select

End If



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



If Site = "Stone" Then

    HDelta = 1

    Sheets("Stream Chemistry").Select

    Sheets("Stream Chemistry").Range("c38").Select: DataCount = ActiveCell.Value

    Sheets("Stream Chemistry").Range("b40").Select

End If



If Site = "Vet's" Then

    HDelta = 1

    Sheets("Stream Chemistry").Select

    Sheets("Stream Chemistry").Range("f38").Select: DataCount = ActiveCell.Value

    Sheets("Stream Chemistry").Range("e40").Select

End If



If Site = "Haze" Then

    HDelta = 1

    Sheets("Stream Chemistry").Select

    Sheets("Stream Chemistry").Range("i38").Select: DataCount = ActiveCell.Value

    Sheets("Stream Chemistry").Range("h40").Select

End If



If Site = "Carter" Then

    HDelta = 1

    Sheets("Stream Chemistry").Select

    Sheets("Stream Chemistry").Range("l38").Select: DataCount = ActiveCell.Value

    Sheets("Stream Chemistry").Range("k40").Select

End If



If Site = "Pioneer" Then

    HDelta = 1

    Sheets("Stream Chemistry").Select

    Sheets("Stream Chemistry").Range("o38").Select: DataCount = ActiveCell.Value

    Sheets("Stream Chemistry").Range("n40").Select

End If



If Site = "USGS" Then

    HDelta = 1

    Sheets("Stream Chemistry").Select

    Sheets("Stream Chemistry").Range("r38").Select: DataCount = ActiveCell.Value

    Sheets("Stream Chemistry").Range("q40").Select

End If



If Site = "NB Ind Hill" Then

    HDelta = 1

    Sheets("Stream Chemistry").Select

    Sheets("Stream Chemistry").Range("u38").Select: DataCount = ActiveCell.Value

    Sheets("Stream Chemistry").Range("t40").Select

End If



If Site = "NB Dead" Then

    HDelta = 1

    Sheets("Stream Chemistry").Select

    Sheets("Stream Chemistry").Range("x38").Select: DataCount = ActiveCell.Value

    Sheets("Stream Chemistry").Range("w40").Select

End If



If Site = "NB Hooker" Then

    HDelta = 1

    Sheets("Stream Chemistry").Select

    Sheets("Stream Chemistry").Range("aa38").Select: DataCount = ActiveCell.Value

    Sheets("Stream Chemistry").Range("z40").Select

End If



If Site = "M22" Then

    HDelta = 1

    Sheets("Stream Chemistry").Select

    Sheets("Stream Chemistry").Range("ad38").Select: DataCount = ActiveCell.Value

    Sheets("Stream Chemistry").Range("ac40").Select

End If



If Site = "BC Old Res" Then

    HDelta = 1

    Sheets("Stream Chemistry").Select

    Sheets("Stream Chemistry").Range("ag38").Select: DataCount = ActiveCell.Value

    Sheets("Stream Chemistry").Range("af40").Select

End If



If Site = "Collision" Then

    HDelta = 1

    Sheets("Stream Chemistry").Select

    Sheets("Stream Chemistry").Range("aj38").Select: DataCount = ActiveCell.Value

    Sheets("Stream Chemistry").Range("ai40").Select

End If



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



For i = 1 To DataCount

    

    If Month(ActiveCell.Value) = 1 And Year(ActiveCell.Value) >= StartYear And Year(ActiveCell.Value) <= EndYear Then

        ActiveCell.Offset(0, HDelta).Select

        M1(Count1) = ActiveCell.Value

        Count1 = Count1 + 1

        ActiveCell.Offset(1, -HDelta).Select

    End If



    If Month(ActiveCell.Value) = 2 And Year(ActiveCell.Value) >= StartYear And Year(ActiveCell.Value) <= EndYear Then

        ActiveCell.Offset(0, HDelta).Select

        M2(Count2) = ActiveCell.Value

        Count2 = Count2 + 1

        ActiveCell.Offset(0, -HDelta).Select

    End If



    If Month(ActiveCell.Value) = 3 And Year(ActiveCell.Value) >= StartYear And Year(ActiveCell.Value) <= EndYear Then

        ActiveCell.Offset(0, HDelta).Select

        M3(Count3) = ActiveCell.Value

        Count3 = Count3 + 1

        ActiveCell.Offset(0, -HDelta).Select

    End If

    

    If Month(ActiveCell.Value) = HDelta And Year(ActiveCell.Value) >= StartYear And Year(ActiveCell.Value) <= EndYear Then

        ActiveCell.Offset(0, HDelta).Select

        M4(Count4) = ActiveCell.Value

        Count4 = Count4 + 1

        ActiveCell.Offset(0, -HDelta).Select

    End If



    If Month(ActiveCell.Value) = 5 And Year(ActiveCell.Value) >= StartYear And Year(ActiveCell.Value) <= EndYear Then

        ActiveCell.Offset(0, HDelta).Select

        M5(Count5) = ActiveCell.Value

        Count5 = Count5 + 1

        ActiveCell.Offset(0, -HDelta).Select

    End If

    

    If Month(ActiveCell.Value) = 6 And Year(ActiveCell.Value) >= StartYear And Year(ActiveCell.Value) <= EndYear Then

        ActiveCell.Offset(0, HDelta).Select

        M6(Count6) = ActiveCell.Value

        Count6 = Count6 + 1

        ActiveCell.Offset(0, -HDelta).Select

    End If



    If Month(ActiveCell.Value) = 7 And Year(ActiveCell.Value) >= StartYear And Year(ActiveCell.Value) <= EndYear Then

        ActiveCell.Offset(0, HDelta).Select

        M7(Count7) = ActiveCell.Value

        Count7 = Count7 + 1

        ActiveCell.Offset(0, -HDelta).Select

    End If



    If Month(ActiveCell.Value) = 8 And Year(ActiveCell.Value) >= StartYear And Year(ActiveCell.Value) <= EndYear Then

        ActiveCell.Offset(0, HDelta).Select

        M8(Count8) = ActiveCell.Value

        Count8 = Count8 + 1

        ActiveCell.Offset(0, -HDelta).Select

    End If



    If Month(ActiveCell.Value) = 9 And Year(ActiveCell.Value) >= StartYear And Year(ActiveCell.Value) <= EndYear Then

        ActiveCell.Offset(0, HDelta).Select

        M9(Count9) = ActiveCell.Value

        Count9 = Count9 + 1

        ActiveCell.Offset(0, -HDelta).Select

    End If



    If Month(ActiveCell.Value) = 10 And Year(ActiveCell.Value) >= StartYear And Year(ActiveCell.Value) <= EndYear Then

        ActiveCell.Offset(0, HDelta).Select

        M10(Count10) = ActiveCell.Value

        Count10 = Count10 + 1

        ActiveCell.Offset(0, -HDelta).Select

    End If



    If Month(ActiveCell.Value) = 11 And Year(ActiveCell.Value) >= StartYear And Year(ActiveCell.Value) <= EndYear Then

        ActiveCell.Offset(0, HDelta).Select

        M11(Count11) = ActiveCell.Value

        Count11 = Count11 + 1

        ActiveCell.Offset(0, -HDelta).Select

    End If

    

    If Month(ActiveCell.Value) = 12 And Year(ActiveCell.Value) >= StartYear And Year(ActiveCell.Value) <= EndYear Then

        ActiveCell.Offset(0, HDelta).Select

        M12(Count12) = ActiveCell.Value

        Count12 = Count12 + 1

        ActiveCell.Offset(0, -HDelta).Select

    End If

 

If Year(ActiveCell.Value) = CompareYear Then

        CompareDate(CompareCount) = ActiveCell.Value

        ActiveCell.Offset(0, HDelta).Select

        CompareValue(CompareCount) = ActiveCell.Value

        CompareCount = CompareCount + 1

        ActiveCell.Offset(0, -HDelta).Select

End If



ActiveCell.Offset(1, 0).Select



Next i





'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx

Sheets("Moving Average").Select

Sheets("Moving Average").Range("y41").Select

For i = 1 To Count1 - 1

    ActiveCell.Value = M1(i)

    ActiveCell.Offset(1, 0).Select

Next i



Sheets("Moving Average").Range("z41").Select

For i = 1 To Count2 - 1

    ActiveCell.Value = M2(i)

    ActiveCell.Offset(1, 0).Select

Next i



Sheets("Moving Average").Range("aa41").Select

For i = 1 To Count3 - 1

    ActiveCell.Value = M3(i)

    ActiveCell.Offset(1, 0).Select

Next i



Sheets("Moving Average").Range("ab41").Select

For i = 1 To Count4 - 1

    ActiveCell.Value = M4(i)

    ActiveCell.Offset(1, 0).Select

Next i



Sheets("Moving Average").Range("ac41").Select

For i = 1 To Count5 - 1

    ActiveCell.Value = M5(i)

    ActiveCell.Offset(1, 0).Select

Next i



Sheets("Moving Average").Range("ad41").Select

For i = 1 To Count6 - 1

    ActiveCell.Value = M6(i)

    ActiveCell.Offset(1, 0).Select

Next i



Sheets("Moving Average").Range("ae41").Select

For i = 1 To Count7 - 1

    ActiveCell.Value = M7(i)

    ActiveCell.Offset(1, 0).Select

Next i



Sheets("Moving Average").Range("af41").Select

For i = 1 To Count8 - 1

    ActiveCell.Value = M8(i)

    ActiveCell.Offset(1, 0).Select

Next i



Sheets("Moving Average").Range("ag41").Select

For i = 1 To Count9 - 1

    ActiveCell.Value = M9(i)

    ActiveCell.Offset(1, 0).Select

Next i



Sheets("Moving Average").Range("ah41").Select

For i = 1 To Count10 - 1

    ActiveCell.Value = M10(i)

    ActiveCell.Offset(1, 0).Select

Next i



Sheets("Moving Average").Range("ai41").Select

For i = 1 To Count11 - 1

    ActiveCell.Value = M11(i)

    ActiveCell.Offset(1, 0).Select

Next i



Sheets("Moving Average").Range("aj41").Select

For i = 1 To Count12 - 1

    ActiveCell.Value = M12(i)

    ActiveCell.Offset(1, 0).Select

Next i



Sheets("Moving Average").Range("t41").Select:

For i = 1 To CompareCount - 1

        ActiveCell.Value = CompareDate(i)

        ActiveCell.Offset(0, 2).Select

        ActiveCell.Value = CompareValue(i)

        CompareCount = CompareCount + 1

        ActiveCell.Offset(0, -2).Select

ActiveCell.Offset(1, 0).Select



Next i



'Sheets("Moving Average").Range("d39").Select  'find maxiumum

 Sheets("Moving Average").Range("L10").Select  'find maxiumum

     MaxValue = ActiveCell.Value

     

     ActiveSheet.ChartObjects("Chart 30").Activate

     ActiveChart.ChartArea.Select

  

     If MaxValue = 16 Then

        With ActiveChart.Axes(xlValue)

            .MinimumScale = 0

            .MaximumScale = MaxValue

            .MajorUnit = 4

        End With

     End If

       

     If MaxValue = 24 Then

        With ActiveChart.Axes(xlValue)

            .MinimumScale = 0

            .MaximumScale = MaxValue

            .MajorUnit = 4

        End With

     End If

     

     If MaxValue = 48 Then

        With ActiveChart.Axes(xlValue)

            .MinimumScale = 0

            .MaximumScale = MaxValue

            .MajorUnit = 8

        End With

     End If

     

     If MaxValue = 96 Then

        With ActiveChart.Axes(xlValue)

            .MinimumScale = 0

            .MaximumScale = MaxValue

            .MajorUnit = 16

        End With

     End If



If Site = "Secchi" Then Sheets("Moving Average").Range("ap9").Select: ActiveCell.Value = "Secchi  Feet"        'chart title

If Site <> "Secchi" Then Sheets("Moving Average").Range("ap9").Select: ActiveCell.Value = "Toral P    mg/m3"        'chart title





Sheets("Moving Average").Range("ap5").Select: ActiveCell.Value = Site & "  " + CStr(StartYear) & " to " + CStr(EndYear)    'define chart title

Sheets("Moving Average").Range("ap6").Select: ActiveCell.Value = "Compared to " + CStr(CompareYear)    'define chart title



Sheets("Moving Average").Range("n3").Select



End Sub



Private Sub worksheet_Activate()

Worksheets("Moving Average").Activate

    Sheets("Moving Average").Select

    TextBox1.Visible = False

    CommandButton3.Caption = "Open"

End Sub

