VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Rectangle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("SteelShapes.Shapes")
Option Explicit

Private Type TRectangle
    Base As Double
    Height As Double
End Type

Private This As TRectangle

Private Const BadDimensionErrorMessage As String = "Rectangle dimension must be greater than zero."
Private Const ClassName As String = "Rectangle"

Public Property Get Base() As Double
    Base = This.Base
End Property

Public Property Let Base(value As Double)
    If value <= 0 Then
        DimensionErrors.ThrowBadDimensionError BadDimensionErrorMessage, ClassName
    End If
    
    This.Base = value
End Property

Public Property Get Height() As Double
    Height = This.Height
End Property

Public Property Let Height(value As Double)
    If value <= 0 Then
        DimensionErrors.ThrowBadDimensionError BadDimensionErrorMessage, ClassName
    End If
    
    This.Height = value
End Property

'@Description "Returns the cross-sectional area"
Public Function Area() As Double
Attribute Area.VB_Description = "Returns the cross-sectional area"
    Area = This.Base * This.Height
End Function

'@Description "Returns the moment of inertia about a horizontal axis that intersects the centroid of the cross-section"
Public Function Ix() As Double
Attribute Ix.VB_Description = "Returns the moment of inertia about a horizontal axis that intersects the centroid of the cross-section"
    Ix = (1 / 12#) * This.Base * This.Height ^ 3
End Function

'@Description "Returns the moment of inertia about a vertical axis that intersects the centroid of the cross-section"
Public Function Iy() As Double
Attribute Iy.VB_Description = "Returns the moment of inertia about a vertical axis that intersects the centroid of the cross-section"
    Iy = (1 / 12#) * This.Base ^ 3 * This.Height
End Function

'@Description "Returns the radius of gyration about a horizontal axis that intersects the centroid of the cross-section"
Public Function rx() As Double
Attribute rx.VB_Description = "Returns the radius of gyration about a horizontal axis that intersects the centroid of the cross-section"
    rx = Math.Sqr(Me.Ix / Me.Area)
End Function

'@Description "Returns the radius of gyration about a vertical axis that intersects the centroid of the cross-section"
Public Function ry() As Double
Attribute ry.VB_Description = "Returns the radius of gyration about a vertical axis that intersects the centroid of the cross-section"
    ry = Math.Sqr(Me.Iy / Me.Area)
End Function

'@Description "Returns the elastic section modulus about a horizontal axis that intersects the centroid of the cross-section"
Public Function Sx() As Double
Attribute Sx.VB_Description = "Returns the elastic section modulus about a horizontal axis that intersects the centroid of the cross-section"
    Sx = (1 / 6#) * This.Base * This.Height ^ 2
End Function

'@Description "Returns the elastic section modulus about a vertical axis that intersects the centroid of the cross-section"
Public Function Sy() As Double
Attribute Sy.VB_Description = "Returns the elastic section modulus about a vertical axis that intersects the centroid of the cross-section"
    Sy = (1 / 6#) * This.Base ^ 2 * This.Height
End Function

'@Description "Returns the plastic section modulus about a horizontal axis that intersects the centroid of the cross-section"
Public Function Zx() As Double
Attribute Zx.VB_Description = "Returns the plastic section modulus about a horizontal axis that intersects the centroid of the cross-section"
    Zx = (1 / 4#) * This.Base * This.Height ^ 2
End Function

'@Description "Returns the plastic section modulus about a vertical axis that intersects the centroid of the cross-section"
Public Function Zy() As Double
Attribute Zy.VB_Description = "Returns the plastic section modulus about a vertical axis that intersects the centroid of the cross-section"
    Zy = (1 / 4#) * This.Base ^ 2 * This.Height
End Function
