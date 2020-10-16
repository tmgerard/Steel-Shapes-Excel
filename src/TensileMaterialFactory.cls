VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TensileMaterialFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "Rubberduck" ,"Predeclared Class Module"
'@Folder("SteelShapes.Material")
Option Explicit
'@PredeclaredId

Public Function Create(ByVal materialGetter As ITensileMaterialGetter, ByVal materialSpecification As String, _
    ByVal materialGrade As String) As ITensileMaterial

    Dim material As TensileMaterial
    Set material = New TensileMaterial
    material.Create materialGetter, materialSpecification, materialGrade
    Set Create = material
End Function
