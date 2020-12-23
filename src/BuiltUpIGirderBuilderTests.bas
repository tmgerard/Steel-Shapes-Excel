Attribute VB_Name = "BuiltUpIGirderBuilderTests"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests.Members")

#If LateBind Then
    Private Assert As Object
    'Private Fakes As Object
#Else
    Private Assert As AssertClass
    'Private Fakes As FakesProvider
#End If

Private builder As BuiltUpIGirderSectionBuilder

' Top Flange
Private topFlangePlate As PlateMemberSection
Private Const topWidth As Double = 12
Private Const topThickness As Double = 0.5
Private Const topMaterialSpec As String = "ASTM A709"
Private Const topMaterialGrade As String = "HPS 70W"

' Web Plate
Private WebPlate As PlateMemberSection
Private Const webWidth As Double = 60
Private Const webThickness As Double = 0.5
Private Const webMaterialSpec As String = "ASTM A709"
Private Const webMaterialGrade As String = "50W"

' Bottom Flange
Private bottomFlangePlate As PlateMemberSection
Private Const bottomWidth As Double = 24
Private Const bottomThickness As Double = 0.5
Private Const bottomMaterialSpec As String = "ASTM A709"
Private Const bottomMaterialGrade As String = "HPS 70W"

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = New Rubberduck.AssertClass
    'Set Fakes = New Rubberduck.FakesProvider
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    'Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
    Set builder = New BuiltUpIGirderSectionBuilder
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
    Set builder = Nothing
End Sub

'@TestMethod("Object Creation")
Private Sub TestBuild()
    On Error GoTo TestFail
    
    'Arrange:
    Dim girder As BuiltUpIGirderSection
    Set girder = New BuiltUpIGirderSection

    'Act:
    With builder
        .SetTopFlange topWidth, topThickness, topMaterialSpec, topMaterialGrade
        .SetWebPlate webWidth, webThickness, webMaterialSpec, webMaterialGrade
        .SetBottomFlange bottomWidth, bottomThickness, bottomMaterialSpec, bottomMaterialGrade
    End With
    Set girder = builder.Build

    'Assert:
    Assert.IsTrue TypeOf girder Is BuiltUpIGirderSection

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Expected Error")
Private Sub TestNoTopFlangeSet()
    Const ExpectedError As Long = BuiltUpIGirderError.MissingPlateObject
    On Error GoTo TestFail
    
    'Arrange:
    Dim girder As BuiltUpIGirderSection
    Set girder = New BuiltUpIGirderSection

    'Act:
    With builder
        .SetWebPlate webWidth, webThickness, webMaterialSpec, webMaterialGrade
        .SetBottomFlange bottomWidth, bottomThickness, bottomMaterialSpec, bottomMaterialGrade
    End With
    Set girder = builder.Build

Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("Expected Error")
Private Sub TestNoWebPlateSet()
    Const ExpectedError As Long = BuiltUpIGirderError.MissingPlateObject
    On Error GoTo TestFail
    
    'Arrange:
    Dim girder As BuiltUpIGirderSection
    Set girder = New BuiltUpIGirderSection

    'Act:
    With builder
        .SetTopFlange topWidth, topThickness, topMaterialSpec, topMaterialGrade
        .SetBottomFlange bottomWidth, bottomThickness, bottomMaterialSpec, bottomMaterialGrade
    End With
    Set girder = builder.Build

Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("Expected Error")
Private Sub TestNoBottomFlangeSet()
    Const ExpectedError As Long = BuiltUpIGirderError.MissingPlateObject
    On Error GoTo TestFail
    
    'Arrange:
    Dim girder As BuiltUpIGirderSection
    Set girder = New BuiltUpIGirderSection

    'Act:
    With builder
        .SetTopFlange topWidth, topThickness, topMaterialSpec, topMaterialGrade
        .SetWebPlate webWidth, webThickness, webMaterialSpec, webMaterialGrade
    End With
    Set girder = builder.Build

Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub
