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
Implements IBuiltUpIGirderComponent

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

Public Function NominalWeight() As Double
    NominalWeight = Material.Density * Me.Area
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

' IBuiltUpIGirderComponent Implementation

Private Function IBuiltUpIGirderComponent_Area() As Double
    IBuiltUpIGirderComponent_Area = Me.Area
End Function

Private Function IBuiltUpIGirderComponent_Ix() As Double
    IBuiltUpIGirderComponent_Ix = Me.Ix
End Function

Private Function IBuiltUpIGirderComponent_Iy() As Double
    IBuiltUpIGirderComponent_Iy = Me.Iy
End Function

Private Property Set IBuiltUpIGirderComponent_Material(ByVal memberMaterial As ITensileMaterial)
    Set Me.Material = memberMaterial
End Property

Private Property Get IBuiltUpIGirderComponent_Material() As ITensileMaterial
    Set IBuiltUpIGirderComponent_Material = Me.Material
End Property

Private Function IBuiltUpIGirderComponent_NominalWeight() As Double
    IBuiltUpIGirderComponent_NominalWeight = Me.NominalWeight
End Function

Private Function IBuiltUpIGirderComponent_rx() As Double
    IBuiltUpIGirderComponent_rx = Me.rx
End Function

Private Function IBuiltUpIGirderComponent_ry() As Double
    IBuiltUpIGirderComponent_ry = Me.ry
End Function

Private Function IBuiltUpIGirderComponent_Sx() As Double
    IBuiltUpIGirderComponent_Sx = Me.Sx
End Function

Private Function IBuiltUpIGirderComponent_Sy() As Double
    IBuiltUpIGirderComponent_Sy = Me.Sy
End Function

Private Property Let IBuiltUpIGirderComponent_Thickness(ByVal value As Double)
    Me.Thickness = value
End Property

Private Property Get IBuiltUpIGirderComponent_Thickness() As Double
    IBuiltUpIGirderComponent_Thickness = Me.Thickness
End Property

Private Property Let IBuiltUpIGirderComponent_Width(ByVal value As Double)
    Me.Width = value
End Property

Private Property Get IBuiltUpIGirderComponent_Width() As Double
    IBuiltUpIGirderComponent_Width = Me.Width
End Property

Private Function IBuiltUpIGirderComponent_Zx() As Double
    IBuiltUpIGirderComponent_Zx = Me.Zx
End Function

Private Function IBuiltUpIGirderComponent_Zy() As Double
    IBuiltUpIGirderComponent_Zy = Me.Zy
End Function
