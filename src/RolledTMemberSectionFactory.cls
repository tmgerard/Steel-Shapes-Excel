VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RolledTMemberSectionFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("SteelShapes.Members")
Option Explicit
'@PredeclaredId

Private Type TRolledTMemberSectionFactory
    shapeGetter As IRolledShapeGetter
    materialGetter As ITensileMaterialGetter
End Type

Private This As TRolledTMemberSectionFactory

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

Public Function Create(ByVal shapeName As String, _
    ByVal materialSpecification As String, _
    ByVal materialGrade As String) As RolledTMemberSection

    Dim newMember As RolledTMemberSection
    Set newMember = New RolledTMemberSection
    With newMember
        Set .Material = TensileMaterialFactory.Create(This.materialGetter, materialSpecification, materialGrade)
        Set .Section = RolledTCrossSectionFactory.Create(This.shapeGetter, shapeName)
    End With
    
    Set Create = newMember

End Function

' NOTE: Might switch to using an enum to use the desired getter object
Public Sub SetShapeGetter(ByVal shapeGetter As IRolledShapeGetter)
    Set This.shapeGetter = shapeGetter
End Sub

Public Sub SetMaterialGetter(ByVal materialGetter As ITensileMaterialGetter)
    Set This.materialGetter = materialGetter
End Sub






