Attribute VB_Name = "PlateMemberSectionTests"
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

Private plate As PlateMemberSection
Private Const pWidth As Double = 12
Private Const pThickness As Double = 1
Private Const doubleComparePrecision As Integer = 4
Private materialGetter As ITensileMaterialGetter
Private Const materialSpec As String = "ASTM A709"
Private Const materialGrade As String = "50W"

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
    
    Set plate = New PlateMemberSection
    Set materialGetter = New CSVTensileMaterialGetter
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    'Set Fakes = Nothing
    Set plate = Nothing
    Set materialGetter = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
    With plate
        .Width = pWidth
        .Thickness = pThickness
        .Orientation = Horizontal
        Set .Material = TensileMaterialFactory.Create(materialGetter, materialSpec, materialGrade)
    End With
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod("Calculation")
Private Sub TestCalculateIxWithHorizontalOrientation()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As Double = 1#

    'Act:

    'Assert:
    Assert.AreEqual Expected, plate.Ix

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Calculation")
Private Sub TestCalculateIyWithHorizontalOrientation()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As Double = 144#

    'Act:

    'Assert:
    Assert.AreEqual Expected, plate.Iy

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Calculation")
Private Sub TestCalculateIxWithVerticalOrientation()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As Double = 144#

    'Act:
    plate.Orientation = Vertical

    'Assert:
    Assert.AreEqual Expected, plate.Ix

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Calculation")
Private Sub TestCalculateIyWithVerticalOrientation()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As Double = 1#

    'Act:
    plate.Orientation = Vertical

    'Assert:
    Assert.AreEqual Expected, plate.Iy

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Calculation")
Private Sub TestCalculaterxWithHorizontalOrientation()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As Double = 0.2887

    'Act:

    'Assert:
    Assert.IsTrue CompareDoubleRound(Expected, plate.rx, doubleComparePrecision)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Calculation")
Private Sub TestCalculateryWithHorizontalOrientation()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As Double = 3.4641

    'Act:

    'Assert:
    Assert.IsTrue CompareDoubleRound(Expected, plate.ry, doubleComparePrecision)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Calculation")
Private Sub TestCalculaterxWithVerticalOrientation()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As Double = 3.4641

    'Act:
    plate.Orientation = Vertical

    'Assert:
    Assert.IsTrue CompareDoubleRound(Expected, plate.rx, doubleComparePrecision)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Calculation")
Private Sub TestCalculateryWithVerticalOrientation()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As Double = 0.2887

    'Act:
    plate.Orientation = Vertical

    'Assert:
    Assert.IsTrue CompareDoubleRound(Expected, plate.ry, doubleComparePrecision)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Calculation")
Private Sub TestCalculateSxWithHorizontalOrientation()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As Double = 2#

    'Act:

    'Assert:
    Assert.AreEqual Expected, plate.Sx

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Calculation")
Private Sub TestCalculateSyWithHorizontalOrientation()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As Double = 24#

    'Act:

    'Assert:
    Assert.AreEqual Expected, plate.Sy

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Calculation")
Private Sub TestCalculateSxWithVerticalOrientation()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As Double = 24#

    'Act:
    plate.Orientation = Vertical

    'Assert:
    Assert.AreEqual Expected, plate.Sx

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Calculation")
Private Sub TestCalculateSyWithVerticalOrientation()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As Double = 2#

    'Act:
    plate.Orientation = Vertical

    'Assert:
    Assert.AreEqual Expected, plate.Sy

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Calculation")
Private Sub TestCalculateZxWithHorizontalOrientation()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As Double = 3#

    'Act:

    'Assert:
    Assert.AreEqual Expected, plate.Zx

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Calculation")
Private Sub TestCalculateZyWithHorizontalOrientation()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As Double = 36#

    'Act:

    'Assert:
    Assert.AreEqual Expected, plate.Zy

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Calculation")
Private Sub TestCalculateZxWithVerticalOrientation()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As Double = 36#

    'Act:
    plate.Orientation = Vertical

    'Assert:
    Assert.AreEqual Expected, plate.Zx

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Calculation")
Private Sub TestCalculateZyWithVerticalOrientation()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As Double = 3#

    'Act:
    plate.Orientation = Vertical

    'Assert:
    Assert.AreEqual Expected, plate.Zy

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Calculation")
Private Sub TestCalculateNominalWeight()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As Double = 3.4028

    'Act:
    plate.Orientation = Vertical

    'Assert:
    Assert.IsTrue CompareDoubleRound(Expected, plate.NominalWeight, doubleComparePrecision)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
