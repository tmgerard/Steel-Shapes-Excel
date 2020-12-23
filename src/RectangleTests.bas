Attribute VB_Name = "RectangleTests"
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

Private rect As Rectangle
Private Const rectWidth As Double = 3
Private Const rectHeight As Double = 4
Private Const doubleComparePrecision As Integer = 4

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
    
    Set rect = New Rectangle
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    'Set Fakes = Nothing
    Set rect = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module.
    rect.Base = rectWidth
    rect.Height = rectHeight
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod("Property")
Private Sub TestGetBase()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As Double = rectWidth

    'Act:

    'Assert:
    Assert.AreEqual Expected, rect.Base

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Property")
Private Sub TestGetNewBase()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As Double = 5

    'Act:
    rect.Base = 5

    'Assert:
    Assert.AreEqual Expected, rect.Base

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Expected Error")
Private Sub TestBadBaseDimension()
    Const ExpectedError As Long = DimensionError.BadDimension
    On Error GoTo TestFail
    
    'Arrange:

    'Act:
    rect.Base = -1

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
Private Sub TestGetHeight()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As Double = rectHeight

    'Act:

    'Assert:
    Assert.AreEqual Expected, rect.Height

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Property")
Private Sub TestGetNewHeight()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As Double = 5

    'Act:
    rect.Height = 5

    'Assert:
    Assert.AreEqual Expected, rect.Height

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Expected Error")
Private Sub TestBadHeightDimension()
    Const ExpectedError As Long = DimensionError.BadDimension
    On Error GoTo TestFail
    
    'Arrange:

    'Act:
    rect.Height = -1

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

'@TestMethod("Calculation")
Private Sub TestCalculateIx()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As Double = 16

    'Act:

    'Assert:
    Assert.AreEqual Expected, rect.Ix

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Calculation")
Private Sub TestCalculateIy()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As Double = 9

    'Act:

    'Assert:
    Assert.AreEqual Expected, rect.Iy

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Calculation")
Private Sub TestCalculaterx()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As Double = 1.1547

    'Act:

    'Assert:
    Assert.IsTrue Compare.CompareDoubleRound(Expected, rect.rx, doubleComparePrecision)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Calculation")
Private Sub TestCalculatery()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As Double = 0.866

    'Act:

    'Assert:
    Assert.IsTrue Compare.CompareDoubleRound(Expected, rect.ry, doubleComparePrecision)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Calculation")
Private Sub TestCalculateSx()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As Double = 8

    'Act:

    'Assert:
    Assert.AreEqual Expected, rect.Sx

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Calculation")
Private Sub TestCalculateSy()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As Double = 6

    'Act:

    'Assert:
    Assert.AreEqual Expected, rect.Sy

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Calculation")
Private Sub TestCalculateZx()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As Double = 12

    'Act:

    'Assert:
    Assert.AreEqual Expected, rect.Zx

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Calculation")
Private Sub TestCalculateZy()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As Double = 9

    'Act:

    'Assert:
    Assert.AreEqual Expected, rect.Zy

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
