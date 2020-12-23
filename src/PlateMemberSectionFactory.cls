VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PlateMemberSectionFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("SteelShapes.Members")
Option Explicit
'@PredeclaredId

Private Type TPlateMemberSectionFactory
    materialGetter As ITensileMaterialGetter
End Type

Private This As TPlateMemberSectionFactory

Private Sub Class_Initialize()
    With This
        Set .materialGetter = New CSVTensileMaterialGetter
    End With
End Sub

Private Sub Class_Terminate()
    With This
        Set .materialGetter = Nothing
    End With
End Sub

Public Function Create(ByVal plateWidth As Double, _
    ByVal plateThickness As Double, _
    ByVal materialSpecification As String, _
    ByVal materialGrade As String, _
    Optional ByVal Orientation As PlateOrientation = PlateOrientation.Horizontal) As PlateMemberSection

    Dim newMember As PlateMemberSection
    Set newMember = New PlateMemberSection
    With newMember
        .width = plateWidth
        .Thickness = plateThickness
        .Orientation = Orientation
        Set .Material = TensileMaterialFactory.Create(This.materialGetter, materialSpecification, materialGrade)
    End With
    
    Set Create = newMember

End Function

' NOTE: Might switch to using an enum to use the desired getter object
Public Sub SetMaterialGetter(ByVal materialGetter As ITensileMaterialGetter)
    Set This.materialGetter = materialGetter
End Sub

