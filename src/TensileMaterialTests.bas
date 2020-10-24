Attribute VB_Name = "TensileMaterialTests"
'@IgnoreModule LineLabelNotUsed, EmptyMethod
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests.Materials")

#If LateBind Then
    Private Assert As Object
    'Private Fakes As Object
#Else
    Private Assert As AssertClass
    'Private Fakes As FakesProvider
#End If

Private materialGetter As ITensileMaterialGetter
Private Material As ITensileMaterial

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
    
    Set materialGetter = New TensileMaterialGetterStub
    Set Material = TensileMaterialFactory.Create(materialGetter, "ASTM A709", "50W")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    'Set Fakes = Nothing
    Set materialGetter = Nothing
    Set Material = Nothing
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
Private Sub TestGetName()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As String = "ASTM A709 50W"

    'Act:

    'Assert:
    Assert.AreEqual Expected, Material.Name

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Property")
Private Sub TestGetYieldStrength()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As Double = 50

    'Act:

    'Assert:
    Assert.AreEqual Expected, Material.YieldStrength

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Property")
Private Sub TestGetUltimateStrength()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As Double = 65

    'Act:

    'Assert:
    Assert.AreEqual Expected, Material.UltimateStrength

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
