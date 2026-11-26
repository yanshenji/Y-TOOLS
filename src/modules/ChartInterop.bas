Attribute VB_Name = "ChartInterop"
Option Explicit

Public Function AddHelperLineFinalProxy(ByVal seriesName As String, ByVal yValue As Double, Optional ByVal showSuccessMessage As Boolean = True) As Boolean
    AddHelperLineFinalProxy = AddHelperLineFinal(seriesName, yValue, showSuccessMessage)
End Function
