VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITensileMaterialGetter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Interface
'@Folder("SteelShapes.Data")
Option Explicit

Public Enum MaterialArrayIndex
    specification
    grade
    Fy
    Fu
End Enum

'@Description "Return material values obtained from permanent storage."
Public Function GetMaterial(ByVal specification As String, ByVal grade As String) As Variant()
Attribute GetMaterial.VB_Description = "Return material values obtained from permanent storage."
End Function
