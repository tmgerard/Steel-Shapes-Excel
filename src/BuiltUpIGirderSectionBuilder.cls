VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BuiltUpIGirderSectionBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("SteelShapes.Members")
Option Explicit

Private Type TBuiltUpIGirderSectionBuilder
    girder As BuiltUpIGirderSection
    topInitialized As Boolean
    webInitialized As Boolean
    bottomInitialized As Boolean
End Type

Private This As TBuiltUpIGirderSectionBuilder
Private Const MissingPlateErrorMessage As String = "Plate member section not initialized"
Private Const ClassName As String = "BuiltUpIGirderSectionBuilder"

Private Sub Class_Initialize()
    With This
        Set .girder = New BuiltUpIGirderSection
        .topInitialized = False
        .webInitialized = False
        .bottomInitialized = False
    End With
End Sub

Private Sub Class_Terminate()
    Set This.girder = Nothing
End Sub

Public Function Build() As BuiltUpIGirderSection
    If Not AllGirderElementsCreated Then
        BuiltUpIGirderErrors.ThrowMissingPlateError MissingPlateErrorMessage, ClassName
    End If
    Set Build = This.girder
End Function

Public Sub SetTopFlange(ByVal PlateWIdth As Double, ByVal Thickness As Double, _
    ByVal materialSpec As String, ByVal materialGrade As String)
    
    Dim plate As PlateMemberSection
    Set plate = PlateMemberSectionFactory.Create(PlateWIdth, Thickness, materialSpec, materialGrade)
    plate.Orientation = Horizontal
    This.topInitialized = True
    
    Set This.girder.TopFlange = plate
End Sub

Public Sub SetWebPlate(ByVal PlateWIdth As Double, ByVal Thickness As Double, _
    ByVal materialSpec As String, ByVal materialGrade As String)
    
    Dim plate As PlateMemberSection
    Set plate = PlateMemberSectionFactory.Create(PlateWIdth, Thickness, materialSpec, materialGrade)
    plate.Orientation = Vertical
    This.webInitialized = True
    
    Set This.girder.WebPlate = plate
End Sub

Public Sub SetBottomFlange(ByVal PlateWIdth As Double, ByVal Thickness As Double, _
    ByVal materialSpec As String, ByVal materialGrade As String)
    
    Dim plate As PlateMemberSection
    Set plate = PlateMemberSectionFactory.Create(PlateWIdth, Thickness, materialSpec, materialGrade)
    plate.Orientation = Horizontal
    This.bottomInitialized = True
    
    Set This.girder.BottomFlange = plate
End Sub

Private Function AllGirderElementsCreated() As Boolean
    AllGirderElementsCreated = This.topInitialized And This.webInitialized And This.bottomInitialized
End Function


