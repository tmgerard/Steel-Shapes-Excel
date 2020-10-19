VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TensileMaterialFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("SteelShapes.Material")
Option Explicit
'@PredeclaredId

Public Function Create(ByVal materialGetter As ITensileMaterialGetter, ByVal materialSpecification As String, _
    ByVal materialGrade As String) As ITensileMaterial

    Dim Material As TensileMaterial
    Set Material = New TensileMaterial
    Material.Create materialGetter, materialSpecification, materialGrade
    Set Create = Material
End Function
