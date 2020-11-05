Attribute VB_Name = "RolledChannelMemberSectionTest"
'@IgnoreModule ProcedureNotUsed
'@Folder("Tests.Members")
Option Explicit
Option Private Module

Private Sub Test()

    Dim member As RolledChannelMemberSection
    Set member = RolledCMemberSectionFactory.Create("C15X33.9", "ASTM A709", "50W")
    
    With member.Section
        Debug.Print "Section Name: " & .Name
        Debug.Print "Area: " & .Area
        Debug.Print "Depth: " & .Depth
        Debug.Print "Flange Width: " & .FlangeWidth
    End With
    
    With member.Material
        Debug.Print "Material Name: " & .Name
        Debug.Print "Yield Strength: " & .YieldStrength
    End With

End Sub

