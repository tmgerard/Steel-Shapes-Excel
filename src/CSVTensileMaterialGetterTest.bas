Attribute VB_Name = "CSVTensileMaterialGetterTest"
'@Folder("Tests.Data")
Option Explicit

Private Sub Test()

    Dim getter As ITensileMaterialGetter
    Set getter = New CSVTensileMaterialGetter
    
    Dim matArray() As String
    
    matArray = getter.GetMaterial("ASTM A709", "50W")
    
    Dim item As Variant
    For Each item In matArray
        Debug.Print item
    Next

End Sub
