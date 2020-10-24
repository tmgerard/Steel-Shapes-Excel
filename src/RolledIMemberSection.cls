VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RolledIMemberSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("SteelShapes.Members")
Option Explicit

Private Type TRolledIMemberSection
    Section As IRolledICrossSection
    Material As ITensileMaterial
End Type

Private This As TRolledIMemberSection

Private Sub Class_Terminate()
    With This
        Set .Section = Nothing
        Set .Material = Nothing
    End With
End Sub

Public Property Get Section() As IRolledICrossSection
    Set Section = This.Section
End Property

Public Property Set Section(ByVal memberSection As IRolledICrossSection)
    Set This.Section = memberSection
End Property

Public Property Get Material() As ITensileMaterial
    Set Material = This.Material
End Property

Public Property Set Material(ByVal memberMaterial As ITensileMaterial)
    Set This.Material = memberMaterial
End Property
