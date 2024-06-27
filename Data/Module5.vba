Attribute VB_Name = "Module5"
Sub Macro3()
Attribute Macro3.VB_Description = "horizontal scale"
Attribute Macro3.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro3 Macro
' horizontal scale
'
'
    ActiveSheet.ChartObjects("Chart 3").Activate
    ActiveChart.Axes(xlCategory).Select
    ActiveChart.Axes(xlCategory).MinimumScale = 1990
    ActiveChart.Axes(xlCategory).MaximumScale = 2022
End Sub