Attribute VB_Name = "CSVRolledShapeGetterTest"
'@Folder("Tests.Data")
Option Explicit
Option Private Module

Private Sub Test()

    Dim getter As IRolledShapeGetter
    Set getter = New CSVRolledShapeGetter
    
    #If Not LateBind Then
        Dim dict As Scripting.Dictionary
    #Else
        Dim dict As Object
    #End If
    
    Set dict = getter.GetRolledShape("HP12x53")
    
    Dim key As Variant
    For Each key In dict
        Debug.Print key, dict.Item(key)
    Next

End Sub
