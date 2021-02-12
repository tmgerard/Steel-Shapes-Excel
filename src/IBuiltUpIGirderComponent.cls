VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IBuiltUpIGirderComponent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Interface
'@Folder("SteelShapes.Members")
Option Explicit

Public Property Get Material() As ITensileMaterial
End Property

Public Property Set Material(ByVal memberMaterial As ITensileMaterial)
End Property

Public Property Get Orientation() As PlateOrientation
End Property

Public Property Get Thickness() As Double
End Property

Public Property Let Thickness(ByVal value As Double)
End Property

Public Property Get PlateWIdth() As Double
End Property

Public Property Let PlateWIdth(ByVal value As Double)
End Property

Public Function Area() As Double
End Function

Public Function Ix() As Double
End Function

Public Function Iy() As Double
End Function

Public Function NominalWeight() As Double
End Function

Public Function rx() As Double
End Function

Public Function ry() As Double
End Function

Public Function Sx() As Double
End Function

Public Function Sy() As Double
End Function

Public Function Zx() As Double
End Function

Public Function Zy() As Double
End Function
