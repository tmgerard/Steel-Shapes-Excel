VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RolledIMemberSectionFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("SteelShapes.Members")
Option Explicit
'@PredeclaredId

Private Type TRolledIMemberSectionFactory
    shapeGetter As IRolledShapeGetter
    materialGetter As ITensileMaterialGetter
End Type

Private This As TRolledIMemberSectionFactory

Private Sub Class_Initialize()
    With This
        Set .shapeGetter = New CSVRolledShapeGetter
        Set .materialGetter = New CSVTensileMaterialGetter
    End With
End Sub

Private Sub Class_Terminate()
    With This
        Set .shapeGetter = Nothing
        Set .materialGetter = Nothing
    End With
End Sub

Public Function Create(ByRef shapeName As String, _
    ByRef materialSpecification As String, _
    ByRef materialGrade As String) As RolledIMemberSection

    Dim newMember As RolledIMemberSection
    Set newMember = New RolledIMemberSection
    With newMember
        Set .Material = TensileMaterialFactory.Create(This.materialGetter, materialSpecification, materialGrade)
        Set .Section = RolledICrossSectionFactory.Create(This.shapeGetter, shapeName)
    End With
    
    Set Create = newMember

End Function

' NOTE: Might switch to using an enum to use the desired getter object
Public Property Set SetShapeGetter(ByRef shapeGetter As IRolledShapeGetter)
    Set This.shapeGetter = shapeGetter
End Property

Public Property Set SetMaterialGetter(ByRef materialGetter As ITensileMaterialGetter)
    Set This.materialGetter = materialGetter
End Property
