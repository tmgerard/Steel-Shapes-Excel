VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TensileMaterialGetterStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Tests.Data.Stubs")
Option Explicit
Implements ITensileMaterialGetter

Private Function ITensileMaterialGetter_GetMaterial(ByVal specification As String, ByVal grade As String) As String()
    Dim Material(0 To 4) As String
    Material(0) = "ASTM A709"
    Material(1) = "50W"
    Material(2) = "50"
    Material(3) = "65"
    Material(4) = "0.28356481"
    ITensileMaterialGetter_GetMaterial = Material
End Function
