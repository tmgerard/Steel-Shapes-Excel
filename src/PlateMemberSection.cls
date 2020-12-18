VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PlateMemberSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("SteelShapes.Members")
Option Explicit

' Plate orientation used as a flag to output
' the correct plate properties for the plates
' positioning within a system
Public Enum PlateOrientation
    Horizontal
    Vertical
End Enum

Private Type TPlateMemberSection
    Shape As Rectangle
    Material As ITensileMaterial
    Orientation As PlateOrientation
End Type

Private This As TPlateMemberSection

Private Sub Class_Initialize()
    Set This.Shape = New Rectangle
    This.Orientation = Horizontal
End Sub

Private Sub Class_Terminate()
    Set This.Shape = Nothing
    Set This.Material = Nothing
End Sub

Public Property Get Material() As ITensileMaterial
    Set Material = This.Material
End Property

Public Property Set Material(ByVal memberMaterial As ITensileMaterial)
    Set This.Material = memberMaterial
End Property

Public Property Get Orientation() As PlateOrientation
    Orientation = This.Orientation
End Property

Public Property Let Orientation(ByVal value As PlateOrientation)
    This.Orientation = value
End Property

Public Property Get Thickness() As Double
    Thickness = This.Shape.Height
End Property

Public Property Let Thickness(ByVal value As Double)
    This.Shape.Height = value
End Property

Public Property Get Width() As Double
    Width = This.Shape.Base
End Property

Public Property Let Width(ByVal value As Double)
    This.Shape.Base = value
End Property

Public Function Area() As Double
    Area = This.Shape.Area
End Function

Public Function Ix() As Double
    If This.Orientation = Horizontal Then
        Ix = This.Shape.Ix
    Else
        Ix = This.Shape.Iy
    End If
End Function

Public Function Iy() As Double
    If This.Orientation = Horizontal Then
        Iy = This.Shape.Iy
    Else
        Iy = This.Shape.Ix
    End If
End Function

Public Function rx() As Double
    If This.Orientation = Horizontal Then
        rx = This.Shape.rx
    Else
        rx = This.Shape.ry
    End If
End Function

Public Function ry() As Double
    If This.Orientation = Horizontal Then
        ry = This.Shape.ry
    Else
        ry = This.Shape.rx
    End If
End Function

Public Function Sx() As Double
    If This.Orientation = Horizontal Then
        Sx = This.Shape.Sx
    Else
        Sx = This.Shape.Sy
    End If
End Function

Public Function Sy() As Double
    If This.Orientation = Horizontal Then
        Sy = This.Shape.Sy
    Else
        Sy = This.Shape.Sx
    End If
End Function

Public Function Zx() As Double
    If This.Orientation = Horizontal Then
        Zx = This.Shape.Zx
    Else
        Zx = This.Shape.Zy
    End If
End Function

Public Function Zy() As Double
    If This.Orientation = Horizontal Then
        Zy = This.Shape.Zy
    Else
        Zy = This.Shape.Zx
    End If
End Function
