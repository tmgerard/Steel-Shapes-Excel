Attribute VB_Name = "BuiltUpIGirderSectionTests"
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

Private plateGirder As BuiltUpIGirderSection

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

Private materialGetter As ITensileMaterialGetter
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
    Set plateGirder = New BuiltUpIGirderSection
    Set topFlangePlate = New PlateMemberSection
    Set WebPlate = New PlateMemberSection
    Set bottomFlangePlate = New PlateMemberSection
    Set materialGetter = New CSVTensileMaterialGetter
    
    With topFlangePlate
        .PlateWIdth = topWidth
        .Thickness = topThickness
        .Orientation = Horizontal
        Set .Material = TensileMaterialFactory.Create(materialGetter, topMaterialSpec, topMaterialGrade)
    End With
    
    With WebPlate
        .PlateWIdth = webWidth
        .Thickness = webThickness
        .Orientation = Vertical
        Set .Material = TensileMaterialFactory.Create(materialGetter, webMaterialSpec, webMaterialGrade)
    End With
    
    With bottomFlangePlate
        .PlateWIdth = bottomWidth
        .Thickness = bottomThickness
        .Orientation = Horizontal
        Set .Material = TensileMaterialFactory.Create(materialGetter, bottomMaterialSpec, bottomMaterialGrade)
    End With
    
    With plateGirder
        Set .TopFlange = topFlangePlate
        Set .WebPlate = WebPlate
        Set .BottomFlange = bottomFlangePlate
    End With
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
    Set plateGirder = Nothing
    Set topFlangePlate = Nothing
    Set WebPlate = Nothing
    Set bottomFlangePlate = Nothing
End Sub

'@TestMethod("Calculation")
Private Sub TestCalculateArea()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As Double = 48#

    'Act:

    'Assert:
    Assert.AreEqual Expected, plateGirder.Area

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Calculation")
Private Sub TestCalculateDepth()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As Double = 61#

    'Act:

    'Assert:
    Assert.AreEqual Expected, plateGirder.Depth

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Calculation")
Private Sub TestCalculateFlangeCentroidToCentroidDistance()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As Double = 60.5

    'Act:

    'Assert:
    Assert.AreEqual Expected, plateGirder.FlangeCentroidToCentroid

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Calculation")
Private Sub TestCalculateWarpingConstant()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As Double = 234256#

    'Act:

    'Assert:
    Assert.AreEqual Expected, plateGirder.Cw

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Calculation")
Private Sub TestCalculateBaseToCentroid()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As Double = 26.7188

    'Act:

    'Assert:
    Assert.IsTrue CompareDoubleRound(Expected, plateGirder.ToCentroid, doubleComparePrecision)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Calculation")
Private Sub TestCalculateIx()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As Double = 24785.2031

    'Act:

    'Assert:
    Assert.IsTrue CompareDoubleRound(Expected, plateGirder.Ix, doubleComparePrecision)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Calculation")
Private Sub TestCalculateIy()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As Double = 648.625

    'Act:

    'Assert:
    Assert.IsTrue CompareDoubleRound(Expected, plateGirder.Iy, doubleComparePrecision)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Calculation")
Private Sub TestCalculateNominalWeight()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As Double = 13.6111

    'Act:

    'Assert:
    Assert.IsTrue CompareDoubleRound(Expected, plateGirder.NominalWeight, doubleComparePrecision)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Calculation")
Private Sub TestCalculateSx()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As Double = 722.9959

    'Act:

    'Assert:
    Assert.IsTrue CompareDoubleRound(Expected, plateGirder.Sx, doubleComparePrecision)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Calculation")
Private Sub TestCalculateSy()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As Double = 54.0521

    'Act:

    'Assert:
    Assert.IsTrue CompareDoubleRound(Expected, plateGirder.Sy, doubleComparePrecision)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Calculation")
Private Sub TestCalculaterx()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As Double = 22.7235

    'Act:

    'Assert:
    Assert.IsTrue CompareDoubleRound(Expected, plateGirder.rx, doubleComparePrecision)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Calculation")
Private Sub TestCalculatery()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As Double = 3.676

    'Act:

    'Assert:
    Assert.IsTrue CompareDoubleRound(Expected, plateGirder.ry, doubleComparePrecision)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Expected Error")
Private Sub TestTopPlateBadOrientation()
    Const ExpectedError As Long = BuiltUpIGirderError.BadPlateOrientation
    On Error GoTo TestFail
    
    'Arrange:
    Dim newPlate As PlateMemberSection
    Set newPlate = New PlateMemberSection
    With newPlate
        .Thickness = topThickness
        .PlateWIdth = topWidth
        .Orientation = Vertical ' top plate should always be horizontal
    End With

    'Act:
    Set plateGirder.TopFlange = newPlate

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
Private Sub TestWebPlateBadOrientation()
    Const ExpectedError As Long = BuiltUpIGirderError.BadPlateOrientation
    On Error GoTo TestFail
    
    'Arrange:
    Dim newPlate As PlateMemberSection
    Set newPlate = New PlateMemberSection
    With newPlate
        .Thickness = webThickness
        .PlateWIdth = webWidth
        .Orientation = Horizontal ' web plate should always be vertical
    End With

    'Act:
    Set plateGirder.WebPlate = newPlate

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
Private Sub TestBottomPlateBadOrientation()
    Const ExpectedError As Long = BuiltUpIGirderError.BadPlateOrientation
    On Error GoTo TestFail
    
    'Arrange:
    Dim newPlate As PlateMemberSection
    Set newPlate = New PlateMemberSection
    With newPlate
        .Thickness = bottomThickness
        .PlateWIdth = bottomWidth
        .Orientation = Vertical ' top plate should always be horizontal
    End With

    'Act:
    Set plateGirder.BottomFlange = newPlate

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
