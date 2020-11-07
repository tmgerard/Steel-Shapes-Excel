VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IRolledTCrossSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule UseMeaningfulName
'@Interface
'@Folder("SteelShapes.Shapes")
Option Explicit

'@Description "Area of the rolled T-shape cross-section"
Public Property Get Area() As Double
Attribute Area.VB_Description = "Area of the rolled T-shape cross-section"
End Property

'@Description "Warping constant of rolled T-shape cross-section"
Public Property Get Cw() As Double
Attribute Cw.VB_Description = "Warping constant of rolled T-shape cross-section"
End Property

'@Description "Depth of rolled T-shape"
Public Property Get Depth() As Double
Attribute Depth.VB_Description = "Depth of rolled T-shape"
End Property

'@Description "Thickness of rolled T-shape flanges"
Public Property Get FlangeThickness() As Double
Attribute FlangeThickness.VB_Description = "Thickness of rolled T-shape flanges"
End Property

'@Description "Width of rolled T-shape flanges"
Public Property Get FlangeWidth() As Double
Attribute FlangeWidth.VB_Description = "Width of rolled T-shape flanges"
End Property

'@Description "Retrieve a property of the rolled T-shape"
Public Function GetProperty(ByRef propertyName As String) As Variant
Attribute GetProperty.VB_Description = "Retrieve a property of the rolled T-shape"
End Function

'@Description "Moment of inertia about the x-axis of the cross-section"
Public Property Get Ix() As Double
Attribute Ix.VB_Description = "Moment of inertia about the x-axis of the cross-section"
End Property

'@Description "Moment of inertia about the y-axis of the cross-section"
Public Property Get Iy() As Double
Attribute Iy.VB_Description = "Moment of inertia about the y-axis of the cross-section"
End Property

'@Description "Torsion constant of the T-shape cross-section"
Public Property Get J() As Double
Attribute J.VB_Description = "Torsion constant of the T-shape cross-section"
End Property

'@Description "Name of rolled T-shape cross-section"
Public Property Get Name() As String
Attribute Name.VB_Description = "Name of rolled T-shape cross-section"
End Property

'@Description "Nominal weight of the rolled T-section"
Public Property Get NominalWeight() As Double
Attribute NominalWeight.VB_Description = "Nominal weight of the rolled T-section"
End Property

'@Description "Radius of gyration about the x-axis"
Public Property Get rx() As Double
Attribute rx.VB_Description = "Radius of gyration about the x-axis"
End Property

'@Description "Radius of gyration about the y-axis"
Public Property Get ry() As Double
Attribute ry.VB_Description = "Radius of gyration about the y-axis"
End Property

'@Description "Elastic section modulus about the x-axis"
Public Property Get Sx() As Double
Attribute Sx.VB_Description = "Elastic section modulus about the x-axis"
End Property

'@Description "Elastic section modulus about the y-axis"
Public Property Get Sy() As Double
Attribute Sy.VB_Description = "Elastic section modulus about the y-axis"
End Property

'@Description "Type of rolled T-shape (WT, MT, or ST)"
Public Property Get ShapeType() As String
Attribute ShapeType.VB_Description = "Type of rolled T-shape (WT, MT, or ST)"
End Property

'@Description "Thickness of rolled T-shape web"
Public Property Get WebThickness() As Double
Attribute WebThickness.VB_Description = "Thickness of rolled T-shape web"
End Property

'@Description "Plastic section modulus about the x-axis"
Public Property Get Zx() As Double
Attribute Zx.VB_Description = "Plastic section modulus about the x-axis"
End Property

'@Description "Plastic section modulus about the y-axis"
Public Property Get Zy() As Double
Attribute Zy.VB_Description = "Plastic section modulus about the y-axis"
End Property
