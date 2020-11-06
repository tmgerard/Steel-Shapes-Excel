VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RolledLShapeGetterStub"
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
        .Add ShapePropertyNames.AiscManualLabel, "L4X3X3/8"
        .Add ShapePropertyNames.ShapeType, "L"
        .Add ShapePropertyNames.Area, 2.49
        .Add ShapePropertyNames.WarpingConstant, 0.114
        .Add ShapePropertyNames.LengthShortAngleLeg, 3#
        .Add ShapePropertyNames.LengthLongAngleLeg, 4#
        .Add ShapePropertyNames.AngleLegThickness, 0.375
        .Add ShapePropertyNames.XAxisMomentOfInertia, 3.94
        .Add ShapePropertyNames.YAxisMomentOfInertia, 1.89
        .Add ShapePropertyNames.ZAxisMomentOfInertia, 1#
        .Add ShapePropertyNames.TorsionalMomentOfInertia, 0.123
        .Add ShapePropertyNames.NominalWeight, 8.5
        .Add ShapePropertyNames.XRadiusOfGyration, 1.26
        .Add ShapePropertyNames.YRadiusOfGyration, 0.873
        .Add ShapePropertyNames.ZRadiusOfGyration, 0.636
        .Add ShapePropertyNames.XAxisSectionModulus, 1.44
        .Add ShapePropertyNames.YAxisSectionModulus, 0.851
        .Add ShapePropertyNames.ZAxisSectionModulus, 0.699
        .Add ShapePropertyNames.XAxisPlasticSectionModulus, 2.6
        .Add ShapePropertyNames.YAxisPlasticSectionModulus, 1.52
    End With
    
    Set IRolledShapeGetter_GetRolledShape = properties

End Function


