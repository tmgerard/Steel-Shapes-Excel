VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITensileMaterial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Interface
'@Folder("SteelShapes.Material")
Option Explicit

'@Description "Name of the material standard."
Public Property Get Name() As String
Attribute Name.VB_Description = "Name of the material standard."
End Property

'@Description "Ultimate tensile strength"
Public Property Get UltimateStrength() As Double
Attribute UltimateStrength.VB_Description = "Ultimate tensile strength"
End Property

'@Description "Tensile strength at first yield"
Public Property Get YieldStrength() As Double
Attribute YieldStrength.VB_Description = "Tensile strength at first yield"
End Property
