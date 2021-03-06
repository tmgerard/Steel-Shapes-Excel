VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BuiltUpIGirderSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("SteelShapes.Members")
Option Explicit

Private Type TBuiltUpIGirderSection
    TopFlange As IBuiltUpIGirderComponent
    WebPlate As IBuiltUpIGirderComponent
    BottomFlange As IBuiltUpIGirderComponent
End Type

Private This As TBuiltUpIGirderSection

Private Const BadWebPlateOrientationErrorMessage As String = "Web plate must have a vertical orientation."
Private Const BadFlangePlateOrientationErrorMessage As String = "Flange plate must have a horizontal orientation."
Private Const ClassName As String = "BuiltUpIGirderSection"

Private Sub Class_Terminate()
    With This
        Set .TopFlange = Nothing
        Set .WebPlate = Nothing
        Set .BottomFlange = Nothing
    End With
End Sub

Public Property Get BottomFlange() As IBuiltUpIGirderComponent
    Set BottomFlange = This.BottomFlange
End Property

Public Property Set BottomFlange(ByVal flange As IBuiltUpIGirderComponent)
    If Not flange.Orientation = Horizontal Then
        BuiltUpIGirderErrors.ThrowBadPlateOrientationError BadFlangePlateOrientationErrorMessage, ClassName
    End If
    Set This.BottomFlange = flange
End Property

Public Property Get TopFlange() As IBuiltUpIGirderComponent
    Set TopFlange = This.TopFlange
End Property

Public Property Set TopFlange(ByVal flange As IBuiltUpIGirderComponent)
    If Not flange.Orientation = Horizontal Then
        BuiltUpIGirderErrors.ThrowBadPlateOrientationError BadFlangePlateOrientationErrorMessage, ClassName
    End If
    Set This.TopFlange = flange
End Property

Public Property Get WebPlate() As IBuiltUpIGirderComponent
    Set WebPlate = This.WebPlate
End Property

Public Property Set WebPlate(ByVal plate As IBuiltUpIGirderComponent)
    If Not plate.Orientation = Vertical Then
        BuiltUpIGirderErrors.ThrowBadPlateOrientationError BadWebPlateOrientationErrorMessage, ClassName
    End If
    Set This.WebPlate = plate
End Property

Public Function Area() As Double
    With This
        Area = .BottomFlange.Area + .TopFlange.Area + .WebPlate.Area
    End With
End Function

' Equation adapted from "Roark's Formulas for Stress and Strain 8th Edition"
' Table 10.2 - Formulas for Torsional Properties and Stresses in Thin-Walled Cross Sections
' Cross Section Reference No. 7
Public Function Cw() As Double
    With This
        Dim numerator As Double
        numerator = Me.FlangeCentroidToCentroid ^ 2 _
            * .TopFlange.Thickness * .BottomFlange.Thickness _
            * .TopFlange.PlateWIdth ^ 3 * .BottomFlange.PlateWIdth ^ 3
            
        Dim denominator As Double
        denominator = 12 * (.TopFlange.Thickness * .TopFlange.PlateWIdth ^ 3 _
            + .BottomFlange.Thickness * .BottomFlange.PlateWIdth ^ 3)
    End With
    
    Cw = numerator / denominator
End Function

Public Function Depth() As Double
    With This
        Depth = .BottomFlange.Thickness + .TopFlange.Thickness + .WebPlate.PlateWIdth
    End With
End Function

Public Function FlangeCentroidToCentroid() As Double
    With This
        FlangeCentroidToCentroid = .WebPlate.PlateWIdth + 0.5 * (.TopFlange.Thickness + .BottomFlange.Thickness)
    End With
End Function

Public Function Ix() As Double
    Ix = BottomFlangeIx + WebPlateIx + TopFlangeIx
End Function

Public Function Iy() As Double
    With This
        Iy = .BottomFlange.Iy + .WebPlate.Iy + .TopFlange.Iy
    End With
End Function

' Equation adapted from "Roark's Formulas for Stress and Strain 8th Edition"
' Table 10.2 - Formulas for Torsional Properties and Stresses in Thin-Walled Cross Sections
' Cross Section Reference No. 7
Public Function J() As Double
    With This
        J = (1 / 3#) _
            * (.TopFlange.Thickness ^ 3 * .TopFlange.PlateWIdth _
            + .BottomFlange.Thickness ^ 3 * .BottomFlange.PlateWIdth _
            + .WebPlate.PlateWIdth ^ 3 * .WebPlate.Thickness)
    End With
End Function

Public Function NominalWeight() As Double
    With This
        NominalWeight = .BottomFlange.NominalWeight + .TopFlange.NominalWeight + .WebPlate.NominalWeight
    End With
End Function

Public Function rx() As Double
    rx = Math.Sqr(Me.Ix / Me.Area)
End Function

Public Function ry() As Double
    ry = Math.Sqr(Me.Iy / Area)
End Function

Public Function Sx() As Double
    Sx = Me.Ix / WorksheetFunction.Max(Me.ToCentroid, Me.Depth - Me.ToCentroid)
End Function

Public Function Sy() As Double
    Sy = Me.Iy / (0.5 * WorksheetFunction.Max(This.TopFlange.PlateWIdth, This.BottomFlange.PlateWIdth))
End Function

'@Description "Distance from the bottom of the girder to the centroid of the girder"
Public Function ToCentroid() As Double
Attribute ToCentroid.VB_Description = "Distance from the bottom of the girder to the centroid of the girder"
    ToCentroid = AreaDistanceSum / Me.Area
End Function

Private Function AreaDistanceSum() As Double
    With This
        Dim result As Double
        ' bottom flange component
        result = .BottomFlange.Area * (.BottomFlange.Thickness / 2#)
        ' web component
        result = result + .WebPlate.Area * (.BottomFlange.Thickness + .WebPlate.PlateWIdth / 2#)
        ' top flange component
        result = result + .TopFlange.Area * (Me.Depth - .TopFlange.Thickness / 2#)
    End With
    
    AreaDistanceSum = result
End Function

Private Function BottomFlangeIx() As Double
    With This.BottomFlange
        Dim Distance As Double
        Distance = Me.ToCentroid - 0.5 * .Thickness
        BottomFlangeIx = ParallelAxisTheorem(.Ix, .Area, Distance)
    End With
End Function

Private Function WebPlateIx() As Double
    With This.WebPlate
        Dim Distance As Double
        Distance = Math.Abs(Me.ToCentroid - (This.BottomFlange.Thickness + 0.5 * .PlateWIdth))
        WebPlateIx = ParallelAxisTheorem(.Ix, .Area, Distance)
    End With
End Function

Private Function TopFlangeIx() As Double
    With This.TopFlange
        Dim Distance As Double
        Distance = (Me.Depth - 0.5 * .Thickness) - Me.ToCentroid
        TopFlangeIx = ParallelAxisTheorem(.Ix, .Area, Distance)
    End With
End Function
