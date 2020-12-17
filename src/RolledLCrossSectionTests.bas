Attribute VB_Name = "RolledLCrossSectionTests"
'@IgnoreModule VariableNotUsed, EmptyMethod, LineLabelNotUsed
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests.Shapes")

#If LateBind Then
    Private Assert As Object
    'Private Fakes As Object
#Else
    Private Assert As AssertClass
    'Private Fakes As FakesProvider
#End If

Private shapeGetter As IRolledShapeGetter
Private interfaceShape As IRolledLCrossSection

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    #If LateBind Then
        Set Assert = CreateObject("Rubberduck.AssertClass")
        'Set Fakes = CreateObject("Rubberduck.FakesProvider")
    #Else
        Set Assert = New AssertClass
        'Set Fakes = New FakesProvider
    #End If
    
    Set shapeGetter = New RolledLShapeGetterStub
    
    ' the concrete implementation exposes the create method
    ' a factory will be used to create sections
    Dim Shape As RolledLCrossSection
    Set Shape = New RolledLCrossSection
    Shape.Create shapeGetter, "L4X3X3/8"
    
    ' the interface exposes the properties
    Set interfaceShape = Shape
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    ' Set Fakes = Nothing
    Set shapeGetter = Nothing
    Set interfaceShape = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod("Property")
Private Sub TestGetArea()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As Double = 2.49

    'Act:

    'Assert:
    Assert.AreEqual Expected, interfaceShape.Area

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Property")
Private Sub TestGetWarpingConstant()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As Double = 0.114

    'Act:

    'Assert:
    Assert.AreEqual Expected, interfaceShape.Cw

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Property")
Private Sub TestGetShortLegLength()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As Double = 3#

    'Act:

    'Assert:
    Assert.AreEqual Expected, interfaceShape.LengthShortLeg

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Property")
Private Sub TestGetLongLegLength()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As Double = 4#

    'Act:

    'Assert:
    Assert.AreEqual Expected, interfaceShape.LengthLongLeg

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Property")
Private Sub TestGetThickness()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As Double = 0.375

    'Act:

    'Assert:
    Assert.AreEqual Expected, interfaceShape.Thickness

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Property")
Private Sub TestGetProperty()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As Double = 0.375

    'Act:

    'Assert:
    Assert.AreEqual Expected, interfaceShape.GetProperty(ShapePropertyNames.AngleLegThickness)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Expected Error")
Private Sub TestGetPropertyKeyDoesNotExist()
    Const ExpectedError As Long = DataError.PropertyNotFound
    On Error GoTo TestFail
    
    'Arrange:
    Dim property As Variant
    
    'Act:
    '@Ignore AssignmentNotUsed
    property = interfaceShape.GetProperty("z")

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

'@TestMethod("Property")
Private Sub TestGetIx()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As Double = 3.94

    'Act:

    'Assert:
    Assert.AreEqual Expected, interfaceShape.Ix

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Property")
Private Sub TestGetIy()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As Double = 1.89

    'Act:

    'Assert:
    Assert.AreEqual Expected, interfaceShape.Iy

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Property")
Private Sub TestGetIz()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As Double = 1#

    'Act:

    'Assert:
    Assert.AreEqual Expected, interfaceShape.Iz

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Property")
Private Sub TestGetJ()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As Double = 0.123

    'Act:

    'Assert:
    Assert.AreEqual Expected, interfaceShape.J

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Property")
Private Sub TestGetName()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As String = "L4X3X3/8"

    'Act:

    'Assert:
    Assert.AreEqual Expected, interfaceShape.Name

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Property")
Private Sub TestGetNominalWeight()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As Double = 8.5

    'Act:

    'Assert:
    Assert.AreEqual Expected, interfaceShape.NominalWeight

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Property")
Private Sub TestGetXRadiusOfGyration()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As Double = 1.26

    'Act:

    'Assert:
    Assert.AreEqual Expected, interfaceShape.rx

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Property")
Private Sub TestGetYRadiusOfGyration()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As Double = 0.873

    'Act:

    'Assert:
    Assert.AreEqual Expected, interfaceShape.ry

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Property")
Private Sub TestGetZRadiusOfGyration()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As Double = 0.636

    'Act:

    'Assert:
    Assert.AreEqual Expected, interfaceShape.rz

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Property")
Private Sub TestGetShapeType()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As String = "L"

    'Act:

    'Assert:
    Assert.AreEqual Expected, interfaceShape.ShapeType

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Property")
Private Sub TestGetSx()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As Double = 1.44

    'Act:

    'Assert:
    Assert.AreEqual Expected, interfaceShape.Sx

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Property")
Private Sub TestGetSz()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As Double = 0.699

    'Act:

    'Assert:
    Assert.AreEqual Expected, interfaceShape.Sz

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Property")
Private Sub TestGetSy()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As Double = 0.851

    'Act:

    'Assert:
    Assert.AreEqual Expected, interfaceShape.Sy

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Property")
Private Sub TestGetZx()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As Double = 2.6

    'Act:

    'Assert:
    Assert.AreEqual Expected, interfaceShape.Zx

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Property")
Private Sub TestGetZy()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As Double = 1.52

    'Act:

    'Assert:
    Assert.AreEqual Expected, interfaceShape.Zy

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


