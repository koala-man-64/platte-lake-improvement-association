Attribute VB_Name = "Sheet18"

Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"

Attribute VB_GlobalNameSpace = False

Attribute VB_Creatable = False

Attribute VB_PredeclaredId = True

Attribute VB_Exposed = True

Attribute VB_TemplateDerived = False

Attribute VB_Customizable = True

Attribute VB_Control = "TextBox1, 4, 0, MSForms, TextBox"

Attribute VB_Control = "CommandButton2, 3, 1, MSForms, CommandButton"

Attribute VB_Control = "CommandButton3, 2, 2, MSForms, CommandButton"

Attribute VB_Control = "CommandButton1, 1, 3, MSForms, CommandButton"

Private Sub CommandButton1_Click()

If CommandButton1.Caption = "Open" Then

        CommandButton1.Caption = "Close"

        TextBox1.Visible = True

    Else

        CommandButton1.Caption = "Open"

        TextBox1.Visible = False

    End If

End Sub

Private Sub CommandButton2_Click()



Dim Flow(500) As Double, WeightedFlow(500) As Double, WetFlow(500) As Double, Rain(500) As Double

Dim EventFactor As Double, SumWetflow As Double, AverageDryFlow As Double, AverageFlow As Double, TotalRain As Double, AverageWetFlow

Dim i As Integer, Spikes As Integer, WetFlowCount As Integer, FlowYear As Integer

Dim FlowColumn As String



Application.ScreenUpdating = False



Sheets("Watershed Mass Bal").Select

Sheets("Watershed Mass Bal").Range("n6").Select

FlowYear = ActiveCell.Value



'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx



Sheets("Annual Averages").Select

Sheets("Annual Averages").Range("C" + CStr(FlowYear - 2010 + 48)).Select

LakeTP = ActiveCell.Value

If LakeTP = 0 Then          'check that data are available for selected year ''''''''''''''''''''''''''''''

   response% = MsgBox(CStr(FlowYear) & " data for Lake TP have are incomplete or entered incorrectly.", 64)

   Exit Sub

End If



ActiveCell.Offset(0, 1).Select

Attainment = ActiveCell.Value

If Attainment = 0 Then          'check that data are available for selected year ''''''''''''''''''''''''''''''

   response% = MsgBox(CStr(FlowYear) & " data for % Attainment are incomplete or entered incorrectly.", 64)

   Exit Sub

End If



ActiveCell.Offset(0, 1).Select

SedRelease = ActiveCell.Value

If SedRelease = 0 Then          'check that data are available for selected year ''''''''''''''''''''''''''''''

   response% = MsgBox(CStr(FlowYear) & " data for Sediment Release are incomplete or entered incorrectly.", 64)

   Exit Sub

End If



ActiveCell.Offset(0, 2).Select

StoneTP = ActiveCell.Value

If StoneTP = 0 Then          'check that data are available for selected year ''''''''''''''''''''''''''''''

   response% = MsgBox(CStr(FlowYear) & " data for Stone TP are incomplete or entered incorrectly.", 64)

   Exit Sub

End If



ActiveCell.Offset(0, 1).Select

CarterTP = ActiveCell.Value

If CarterTP = 0 Then          'check that data are available for selected year ''''''''''''''''''''''''''''''

   response% = MsgBox(CStr(FlowYear) & " data for Carter TP are incomplete or entered incorrectly.", 64)

   Exit Sub

End If



ActiveCell.Offset(0, 1).Select

CollisionTP = ActiveCell.Value

If CollisionTP = 0 Then         'check that data are available for selected year ''''''''''''''''''''''''''''''

   response% = MsgBox(CStr(FlowYear) & " data for Collision TP are incomplete or entered incorrectly .", 64)

   Exit Sub

End If



ActiveCell.Offset(0, 1).Select

NBDeadTP = ActiveCell.Value

If NBDeadTP = 0 Then         'check that data are available for selected year ''''''''''''''''''''''''''''''

   response% = MsgBox(CStr(FlowYear) & " data for Deadstream TP are incomplete or entered incorrectly.", 64)

   Exit Sub

End If



ActiveCell.Offset(0, 1).Select

VetsTP = ActiveCell.Value

If VetsTP = 0 Then         'check that data are available for selected year ''''''''''''''''''''''''''''''

   response% = MsgBox(CStr(FlowYear) & " data for Vet's TP are incomplete or entered incorrectly.", 64)

   Exit Sub

End If



ActiveCell.Offset(0, 1).Select

PioneerTP = ActiveCell.Value

If PioneerTP = 0 Then         'check that data are available for selected year ''''''''''''''''''''''''''''''

   response% = MsgBox(CStr(FlowYear) & " data for Pioneer TP are incomplete or entered incorrectly.", 64)

   Exit Sub

End If



ActiveCell.Offset(0, 1).Select

USGSTP = ActiveCell.Value           'm

If USGSTP = 0 Then         'check that data are available for selected year ''''''''''''''''''''''''''''''

   response% = MsgBox(CStr(FlowYear) & " data for USGS TP are incomplete or entered incorrectly.", 64)

   Exit Sub

End If



ActiveCell.Offset(0, 3).Select

BCInFlow = ActiveCell.Value * 1.547 'p  was mgd  now cfs

If BCInFlow = 0 Then         'check that data are available for selected year ''''''''''''''''''''''''''''''

   response% = MsgBox(CStr(FlowYear) & " data for BC InFlow have not been entered or entered incorrectly.", 64)

   Exit Sub

End If



ActiveCell.Offset(0, 1).Select

BCInTP = ActiveCell.Value

If BCInTP = 0 Then         'check that data are available for selected year ''''''''''''''''''''''''''''''

   response% = MsgBox(CStr(FlowYear) & " data for BC TP have not been entered or entered incorrectly.", 64)

   Exit Sub

End If



ActiveCell.Offset(0, 1).Select

BCInLoad = ActiveCell.Value

If BCInLoad = 0 Then         'check that data are available for selected year ''''''''''''''''''''''''''''''

   response% = MsgBox(CStr(FlowYear) & " data for BC Input Load have not been entered or entered incorrectly.", 64)

   Exit Sub

End If



ActiveCell.Offset(0, 1).Select

HatcheryFlow = ActiveCell.Value * 1.547 'was mgs now cfs

If HatcheryFlow = 0 Then         'check that data are available for selected year ''''''''''''''''''''''''''''''

   response% = MsgBox(CStr(FlowYear) & " data for Hatchery Flow have not been entered or entered incorrectly.", 64)

   Exit Sub

End If



ActiveCell.Offset(0, 1).Select

HatcheryTP = ActiveCell.Value

If HatcheryTP = 0 Then         'check that data are available for selected year ''''''''''''''''''''''''''''''

   response% = MsgBox(CStr(FlowYear) & " data for Hatchery TP have not been entered or entered incorrectly.", 64)

   Exit Sub

End If



ActiveCell.Offset(0, 1).Select

HatcheryLoad = ActiveCell.Value

If HatcheryLoad = 0 Then         'check that data are available for selected year ''''''''''''''''''''''''''''''

   response% = MsgBox(CStr(FlowYear) & " data for Hatchery Load have not been entered or entered incorrectly.", 64)

   Exit Sub

End If



ActiveCell.Offset(0, 1).Select

LostFish = ActiveCell.Value     'v

If LostFish = 0 Then         'check that data are available for selected year ''''''''''''''''''''''''''''''

   response% = MsgBox(CStr(FlowYear) & " data for Lost Fish have not been entered or entered incorrectly.", 64)

   Exit Sub

End If



ActiveCell.Offset(0, 6).Select  'ab

RainLoad = ActiveCell.Value

If RainLoad = 0 Then         'check that data are available for selected year ''''''''''''''''''''''''''''''

   response% = MsgBox(CStr(FlowYear) & " data for Atmospheric Load have not been entered or entered incorrectly.", 64)

   Exit Sub

End If



'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx





Sheets("Flow & Rain & TP Comparison").Select

Sheets("Flow & Rain & TP Comparison").Range("o" + CStr(FlowYear - 2010 + 10)).Select

Events = ActiveCell.Value

If Events = 0 Then         'check that data are available for selected year ''''''''''''''''''''''''''''''

   response% = MsgBox(CStr(FlowYear) & " data for Events have not been entered or entered incorrectly.", 64)

   Exit Sub

End If



ActiveCell.Offset(0, 1).Select

EventFlow = ActiveCell.Value

If EventFlow = 0 Then         'check that data are available for selected year ''''''''''''''''''''''''''''''

   response% = MsgBox(CStr(FlowYear) & " data for Event Flow have not been entered or entered incorrectly.", 64)

   Exit Sub

End If



ActiveCell.Offset(0, 1).Select

BaseFlow = ActiveCell.Value

If BaseFlow = 0 Then         'check that data are available for selected year ''''''''''''''''''''''''''''''

   response% = MsgBox(CStr(FlowYear) & " data for Base Flow have not been entered or entered incorrectly.", 64)

   Exit Sub

End If



ActiveCell.Offset(0, 1).Select

USGSFlow = ActiveCell.Value

If USGSFlow = 0 Then          'check that data are available for selected year ''''''''''''''''''''''''''''''

   response% = MsgBox(CStr(FlowYear) & " data for USGS Flow have not been entered or entered incorrectly.", 64)

   Exit Sub

End If





'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx



''''''''''''''''''''''''''''''' all outside values are in, now place



Sheets("Watershed Mass Bal").Select

Sheets("Watershed Mass Bal").Range("k27").Select

ActiveCell.Value = USGSFlow



Sheets("Watershed Mass Bal").Range("z28").Select

ActiveCell.Value = StoneTP



Sheets("Watershed Mass Bal").Range("w33").Select

ActiveCell.Value = BCInTP



Sheets("Watershed Mass Bal").Range("w32").Select

ActiveCell.Value = BCInFlow

ActiveCell.Offset(2, 0).Select

ActiveCell.Value = BCInLoad



Sheets("Watershed Mass Bal").Range("u32").Select

ActiveCell.Value = HatcheryFlow

ActiveCell.Offset(1, 0).Select

ActiveCell.Value = HatcheryTP

ActiveCell.Offset(1, 0).Select

ActiveCell.Value = HatcheryLoad



Sheets("Watershed Mass Bal").Range("t28").Select

ActiveCell.Value = VetsTP

Sheets("Watershed Mass Bal").Range("q33").Select

ActiveCell.Value = CarterTP

Sheets("Watershed Mass Bal").Range("p28").Select

ActiveCell.Value = PioneerTP

Sheets("Watershed Mass Bal").Range("m22").Select

ActiveCell.Value = CollisionTP

Sheets("Watershed Mass Bal").Range("k28").Select

ActiveCell.Value = USGSTP

Sheets("Watershed Mass Bal").Range("i22").Select

ActiveCell.Value = NBDeadTP



'''''''''''''''''''''



Sheets("Watershed Mass Bal").Range("k32").Select

ActiveCell.Value = Events

ActiveCell.Offset(1, 0).Select

ActiveCell.Value = EventFlow

ActiveCell.Offset(1, 0).Select

ActiveCell.Value = BaseFlow



Sheets("Watershed Mass Bal").Range("f22").Select

ActiveCell.Value = LostFish

Sheets("Watershed Mass Bal").Range("f25").Select

ActiveCell.Value = RainLoad

ActiveCell.Offset(1, 0).Select

ActiveCell.Value = SedRelease

Sheets("Watershed Mass Bal").Range("f30").Select

ActiveCell.Value = LakeTP

ActiveCell.Offset(1, 0).Select

ActiveCell.Value = Attainment



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sheets("Watershed Mass Bal").Range("f32").Select

LossRate = ActiveCell.Value

Sheets("Watershed Mass Bal").Range("f29").Select

TotalLoad = ActiveCell.Value

Sheets("Watershed Mass Bal").Range("s21").Select 'del

NPLoad = ActiveCell.Value                           'del

Sheets("Watershed Mass Bal").Range("z29").Select

UpperLoad = ActiveCell.Value



'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx



Sheets("Annual Averages").Select

Sheets("Annual Averages").Range("f" + CStr(FlowYear - 2010 + 48)).Select

ActiveCell.Value = LossRate



Sheets("Annual Averages").Range("w" + CStr(FlowYear - 2010 + 48)).Select

ActiveCell.Value = TotalLoad

ActiveCell.Offset(0, 1).Select

ActiveCell.Value = TotalLoad - UpperLoad - LostFish - RainLoad - HatcheryLoad - SedRelease

ActiveCell.Offset(0, 1).Select

ActiveCell.Value = UpperLoad



'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx

Sheets("Watershed Mass Bal").Select

Sheets("Watershed Mass Bal").Range("n6").Select



End Sub

Private Sub CommandButton3_Click()

Sheets("Main Menu").Select

    Sheets("Main Menu").Range("g11").Select

End Sub

Private Sub worksheet_Activate()

Worksheets("Watershed Mass Bal").Activate

    Sheets("Watershed Mass Bal").Select

    TextBox1.Visible = False

    CommandButton1.Caption = "Open"

End Sub

