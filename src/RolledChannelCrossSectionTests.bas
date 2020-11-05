Attribute VB_Name = "RolledChannelCrossSectionTests"
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
Private interfaceShape As IRolledICrossSection

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
    
    Set shapeGetter = New RolledChannelShapeGetterStub
    
    ' the concrete implementation exposes the create method
    ' a factory will be used to create sections
    Dim shape As RolledICrossSection
    Set shape = New RolledICrossSection
    shape.Create shapeGetter, "C15X33.9"
    
    ' the interface exposes the properties
    Set interfaceShape = shape
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
    Const Expected As Double = 10#

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
    Const Expected As Double = 358

    'Act:

    'Assert:
    Assert.AreEqual Expected, interfaceShape.Cw

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Property")
Private Sub TestGetDepth()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As Double = 15#

    'Act:

    'Assert:
    Assert.AreEqual Expected, interfaceShape.Depth

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Property")
Private Sub TestGetFlangeThickness()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As Double = 0.65

    'Act:

    'Assert:
    Assert.AreEqual Expected, interfaceShape.FlangeThickness

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Property")
Private Sub TestGetFlangeWidth()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As Double = 3.4

    'Act:

    'Assert:
    Assert.AreEqual Expected, interfaceShape.FlangeWidth

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Property")
Private Sub TestGetProperty()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As Double = 3.4

    'Act:

    'Assert:
    Assert.AreEqual Expected, interfaceShape.GetProperty(ShapePropertyNames.FlangeWidth)

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
    Const Expected As Double = 315

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
    Const Expected As Double = 8.07

    'Act:

    'Assert:
    Assert.AreEqual Expected, interfaceShape.Iy

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Property")
Private Sub TestGetJ()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As Double = 1.01

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
    Const Expected As String = "C15X33.9"

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
    Const Expected As Double = 33.9

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
    Const Expected As Double = 5.61

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
    Const Expected As Double = 0.901

    'Act:

    'Assert:
    Assert.AreEqual Expected, interfaceShape.ry

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Property")
Private Sub TestGetShapeType()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As String = "C"

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
    Const Expected As Double = 42#

    'Act:

    'Assert:
    Assert.AreEqual Expected, interfaceShape.Sx

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Property")
Private Sub TestGetSy()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As Double = 3.09

    'Act:

    'Assert:
    Assert.AreEqual Expected, interfaceShape.Sy

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Property")
Private Sub TestGetWebThickness()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As Double = 0.4

    'Act:

    'Assert:
    Assert.AreEqual Expected, interfaceShape.WebThickness

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Property")
Private Sub TestGetZx()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As Double = 50.8

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
    Const Expected As Double = 6.19

    'Act:

    'Assert:
    Assert.AreEqual Expected, interfaceShape.Zy

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

