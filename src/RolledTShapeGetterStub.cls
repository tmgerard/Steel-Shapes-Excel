VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RolledTShapeGetterStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Tests.Data.Stubs")
Option Explicit

Implements IRolledShapeGetter

Private Function IRolledShapeGetter_GetRolledShape(ByVal shapeName As String) As Object

    #If LateBind Then
        Dim properties As Object
        Set properties = CreateObject("Scripting.Dictionary")
    #Else
        Dim properties As Scripting.Dictionary
        Set properties = New Scripting.Dictionary
    #End If
    
    With properties
        .Add ShapePropertyNames.AiscManualLabel, "WT18X80"
        .Add ShapePropertyNames.ShapeType, "WT"
        .Add ShapePropertyNames.Area, 23.5
        .Add ShapePropertyNames.WarpingConstant, 53.6
        .Add ShapePropertyNames.OverallDepth, 18#
        .Add ShapePropertyNames.FlangeThickness, 1.02
        .Add ShapePropertyNames.FlangeWidth, 12#
        .Add ShapePropertyNames.XAxisMomentOfInertia, 740
        .Add ShapePropertyNames.YAxisMomentOfInertia, 147
        .Add ShapePropertyNames.TorsionalMomentOfInertia, 6.17
        .Add ShapePropertyNames.NominalWeight, 80
        .Add ShapePropertyNames.XRadiusOfGyration, 5.61
        .Add ShapePropertyNames.YRadiusOfGyration, 2.5
        .Add ShapePropertyNames.XAxisSectionModulus, 55.8
        .Add ShapePropertyNames.YAxisSectionModulus, 24.8
        .Add ShapePropertyNames.webThickness, 0.65
        .Add ShapePropertyNames.XAxisPlasticSectionModulus, 100
        .Add ShapePropertyNames.YAxisPlasticSectionModulus, 38.6
    End With
    
    Set IRolledShapeGetter_GetRolledShape = properties

End Function

