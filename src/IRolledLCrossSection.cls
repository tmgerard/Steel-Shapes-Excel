VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IRolledLCrossSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule UseMeaningfulName
'@Interface
'@Folder("SteelShapes.Shapes")
Option Explicit

'@Description "Area of the single-angle cross-section"
Public Property Get Area() As Double
Attribute Area.VB_Description = "Area of the single-angle cross-section"
End Property

'@Description "Warping constant of single-angle cross-section"
Public Property Get Cw() As Double
Attribute Cw.VB_Description = "Warping constant of single-angle cross-section"
End Property

'@Description "Retrieve a property of the single-angle cross-section"""
Public Function GetProperty(ByRef propertyName As String) As Variant
Attribute GetProperty.VB_Description = "Retrieve a property of the single-angle cross-section"""
End Function

'@Description "Moment of inertia about the x-axis of the cross-section"
Public Property Get Ix() As Double
Attribute Ix.VB_Description = "Moment of inertia about the x-axis of the cross-section"
End Property

'@Description "Moment of inertia about the y-axis of the cross-section"
Public Property Get Iy() As Double
Attribute Iy.VB_Description = "Moment of inertia about the y-axis of the cross-section"
End Property

'@Description "Moment of interia about the z-axis (principal axis) of the cross-section"
Public Property Get Iz() As Double
Attribute Iz.VB_Description = "Moment of interia about the z-axis (principal axis) of the cross-section"
End Property

'@Description "Torsion constant of single-angle cross-section"
Public Property Get J() As Double
Attribute J.VB_Description = "Torsion constant of single-angle cross-section"
End Property

'@Description "Length of the longer leg in the single-angle cross-section"
Public Property Get LengthLongLeg() As Double
Attribute LengthLongLeg.VB_Description = "Length of the longer leg in the single-angle cross-section"
End Property

'@Description "Length of the shorter leg in the single-angle cross-section"
Public Property Get LengthShortLeg() As Double
Attribute LengthShortLeg.VB_Description = "Length of the shorter leg in the single-angle cross-section"
End Property

'@Description "Name of single-angle cross-section"
Public Property Get Name() As String
Attribute Name.VB_Description = "Name of single-angle cross-section"
End Property

'@Description "Nominal weight of single-angle cross-section"
Public Property Get NominalWeight() As Double
Attribute NominalWeight.VB_Description = "Nominal weight of single-angle cross-section"
End Property

'@Description "Radius of gyration about the x-axis"
Public Property Get rx() As Double
Attribute rx.VB_Description = "Radius of gyration about the x-axis"
End Property

'@Description "Radius of gyration about the y-axis"
Public Property Get ry() As Double
Attribute ry.VB_Description = "Radius of gyration about the y-axis"
End Property

'@Description "Radius of gyration about the z-axis (principal axis)"
Public Property Get rz() As Double
Attribute rz.VB_Description = "Radius of gyration about the z-axis (principal axis)"
End Property

'@Description "Elastic section modulus about the x-axis"
Public Property Get Sx() As Double
Attribute Sx.VB_Description = "Elastic section modulus about the x-axis"
End Property

'@Description "Elastic section modulus about the y-axis"
Public Property Get Sy() As Double
Attribute Sy.VB_Description = "Elastic section modulus about the y-axis"
End Property

'@Description "Elastic section modulus about the z-axis (principal axis)"
Public Property Get Sz() As Double
Attribute Sz.VB_Description = "Elastic section modulus about the z-axis (principal axis)"
End Property

'@Description "Type of rolled shape (L)"
Public Property Get ShapeType() As String
Attribute ShapeType.VB_Description = "Type of rolled shape (L)"
End Property

'@Description "Thickness of angle leg"
Public Property Get Thickness() As Double
Attribute Thickness.VB_Description = "Thickness of angle leg"
End Property

'@Description "Plastic section modulus about the x-axis"
Public Property Get Zx() As Double
Attribute Zx.VB_Description = "Plastic section modulus about the x-axis"
End Property

'@Description "Plastic section modulus about the y-axis"
Public Property Get Zy() As Double
Attribute Zy.VB_Description = "Plastic section modulus about the y-axis"
End Property
