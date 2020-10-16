VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TensileMaterial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("SteelShapes.Material")
Option Explicit
Implements ITensileMaterial

Private Type TTensileMaterial
    grade As String
    specification As String
    UltimateStrength As Double
    YieldStrength As Double
End Type

Private This As TTensileMaterial

Public Sub Create(ByVal materialGetter As ITensileMaterialGetter, ByVal specification As String, _
    ByVal grade As String)
    
    Dim properties() As String
    properties = materialGetter.GetMaterial(specification, grade)
    
    With This
        .specification = properties(MaterialArrayIndex.specification)
        .grade = properties(MaterialArrayIndex.grade)
        .UltimateStrength = CDbl(properties(MaterialArrayIndex.Fu))
        .YieldStrength = CDbl(properties(MaterialArrayIndex.Fy))
    End With
    
End Sub

' ITensileMaterial Implementation

Private Property Get ITensileMaterial_Name() As String
    ITensileMaterial_Name = This.specification & " " & This.grade
End Property

Private Property Get ITensileMaterial_UltimateStrength() As Double
    ITensileMaterial_UltimateStrength = This.UltimateStrength
End Property

Private Property Get ITensileMaterial_YieldStrength() As Double
    ITensileMaterial_YieldStrength = This.YieldStrength
End Property
