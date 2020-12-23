VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RolledChannelShapeGetterStub"
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
        .Add ShapePropertyNames.AiscManualLabel, "C15X33.9"
        .Add ShapePropertyNames.ShapeType, "C"
        .Add ShapePropertyNames.Area, 10#
        .Add ShapePropertyNames.WarpingConstant, 358
        .Add ShapePropertyNames.OverallDepth, 15#
        .Add ShapePropertyNames.FlangeThickness, 0.65
        .Add ShapePropertyNames.FlangeWidth, 3.4
        .Add ShapePropertyNames.XAxisMomentOfInertia, 315
        .Add ShapePropertyNames.YAxisMomentOfInertia, 8.07
        .Add ShapePropertyNames.TorsionalMomentOfInertia, 1.01
        .Add ShapePropertyNames.NominalWeight, 33.9
        .Add ShapePropertyNames.XRadiusOfGyration, 5.61
        .Add ShapePropertyNames.YRadiusOfGyration, 0.901
        .Add ShapePropertyNames.XAxisSectionModulus, 42#
        .Add ShapePropertyNames.YAxisSectionModulus, 3.09
        .Add ShapePropertyNames.webThickness, 0.4
        .Add ShapePropertyNames.XAxisPlasticSectionModulus, 50.8
        .Add ShapePropertyNames.YAxisPlasticSectionModulus, 6.19
    End With
    
    Set IRolledShapeGetter_GetRolledShape = properties

End Function

