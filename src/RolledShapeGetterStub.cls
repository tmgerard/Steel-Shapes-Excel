VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RolledShapeGetterStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Tests.Data.Stubs")
Option Explicit
Implements IRolledShapeGetter

Private Function IRolledShapeGetter_GetRolledShape(shapeName As String) As Object

    #If LateBind Then
        Dim properties As Object
        Set properties = CreateObject("Scripting.Dictionary")
    #Else
        Dim properties As Scripting.Dictionary
        Set properties = New Scripting.Dictionary
    #End If
    
    With properties
        .Add ShapePropertyNames.AiscManualLabel, "W44X335"
        .Add ShapePropertyNames.ShapeType, "W"
        .Add ShapePropertyNames.Area, 98.5
        .Add ShapePropertyNames.WarpingConstant, 535000
        .Add ShapePropertyNames.OverallDepth, 44
        .Add ShapePropertyNames.FlangeThickness, 1.77
        .Add ShapePropertyNames.FlangeWidth, 15.9
        .Add ShapePropertyNames.XAxisMomentOfInertia, 31100
        .Add ShapePropertyNames.YAxisMomentOfInertia, 1200
        .Add ShapePropertyNames.TorsionalMomentOfInertia, 74.7
        .Add ShapePropertyNames.NominalWeight, 335
        .Add ShapePropertyNames.XRadiusOfGyration, 17.8
        .Add ShapePropertyNames.YRadiusOfGyration, 3.49
        .Add ShapePropertyNames.XAxisSectionModulus, 1410
        .Add ShapePropertyNames.YAxisSectionModulus, 150
        .Add ShapePropertyNames.WebThickness, 1.03
        .Add ShapePropertyNames.XAxisPlasticSectionModulus, 1620
        .Add ShapePropertyNames.YAxisPlasticSectionModulus, 236
    End With
    
    Set IRolledShapeGetter_GetRolledShape = properties

End Function
