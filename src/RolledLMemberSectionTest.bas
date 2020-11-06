Attribute VB_Name = "RolledLMemberSectionTest"
'@IgnoreModule ProcedureNotUsed
'@Folder("Tests.Members")
Option Explicit
Option Private Module

Private Sub Test()

    Dim member As RolledLMemberSection
    Set member = RolledLMemberSectionFactory.Create("L4X3X3/8", "ASTM A709", "50W")
    
    With member.Section
        Debug.Print "Section Name: " & .Name
        Debug.Print "Area: " & .Area
        Debug.Print "Long Leg Length: " & .LengthLongLeg
        Debug.Print "Short Leg Length: " & .LengthShortLeg
        Debug.Print "Thickness: " & .Thickness
    End With
    
    With member.Material
        Debug.Print "Material Name: " & .Name
        Debug.Print "Yield Strength: " & .YieldStrength
    End With

End Sub


