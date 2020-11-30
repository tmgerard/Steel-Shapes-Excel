Attribute VB_Name = "Compare"
'@Folder("SteelShapes.General")
Option Explicit

Public Function CompareDoubleRound(ByVal ValueOne As Double, ByVal ValueTwo As Double, _
    Optional ByVal Precision As Long = 8) As Boolean
    CompareDoubleRound = (Math.Round(ValueOne, Precision) = Math.Round(ValueTwo, Precision))
End Function

