VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RolledChannelMemberSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("SteelShapes.Members")
Option Explicit

Private Type TRolledChannelMemberSection
    Section As IRolledChannelCrossSection
    Material As ITensileMaterial
End Type

Private This As TRolledChannelMemberSection

Private Sub Class_Terminate()
    With This
        Set .Section = Nothing
        Set .Material = Nothing
    End With
End Sub

Public Property Get Section() As IRolledChannelCrossSection
    Set Section = This.Section
End Property

Public Property Set Section(ByVal memberSection As IRolledChannelCrossSection)
    Set This.Section = memberSection
End Property

Public Property Get Material() As ITensileMaterial
    Set Material = This.Material
End Property

Public Property Set Material(ByVal memberMaterial As ITensileMaterial)
    Set This.Material = memberMaterial
End Property

