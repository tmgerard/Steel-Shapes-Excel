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

Public Function Area() As Double
    Area = This.Base * This.Height
End Function

Public Function Ix() As Double
    Ix = (1 / 12#) * This.Base * This.Height ^ 3
End Function

Public Function Iy() As Double
    Iy = (1 / 12#) * This.Base ^ 3 * This.Height
End Function

Public Function rx() As Double
    rx = Math.Sqr(Me.Ix / Me.Area)
End Function

Public Function ry() As Double
    ry = Math.Sqr(Me.Iy / Me.Area)
End Function

Public Function Sx() As Double
    Sx = (1 / 6#) * This.Base * This.Height ^ 2
End Function

Public Function Sy() As Double
    Sy = (1 / 6#) * This.Base ^ 2 * This.Height
End Function

Public Function Zx() As Double
    Zx = (1 / 4#) * This.Base * This.Height ^ 2
End Function

Public Function Zy() As Double
    Zy = (1 / 4#) * This.Base ^ 2 * This.Height
End Function
