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
End Type

Private This As TBuiltUpIGirderSectionBuilder
Private Const MissingPlateErrorMessage As String = "Plate member section not initialized"
Private Const ClassName As String = "BuiltUpIGirderSectionBuilder"

Private Sub Class_Initialize()
    Set This.girder = New BuiltUpIGirderSection
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

Public Sub SetTopFlange(ByVal width As Double, ByVal Thickness As Double, _
    ByVal materialSpec As String, ByVal materialGrade As String)
    
    Dim plate As PlateMemberSection
    Set plate = PlateMemberSectionFactory.Create(width, Thickness, materialSpec, materialGrade)
    plate.Orientation = Horizontal
    
    Set This.girder.TopFlange = plate
End Sub

Public Sub SetWebPlate(ByVal width As Double, ByVal Thickness As Double, _
    ByVal materialSpec As String, ByVal materialGrade As String)
    
    Dim plate As PlateMemberSection
    Set plate = PlateMemberSectionFactory.Create(width, Thickness, materialSpec, materialGrade)
    plate.Orientation = Vertical
    
    Set This.girder.TopFlange = plate
End Sub

Public Sub SetBottomFlange(ByVal width As Double, ByVal Thickness As Double, _
    ByVal materialSpec As String, ByVal materialGrade As String)
    
    Dim plate As PlateMemberSection
    Set plate = PlateMemberSectionFactory.Create(width, Thickness, materialSpec, materialGrade)
    plate.Orientation = Horizontal
    
    Set This.girder.TopFlange = plate
End Sub

Private Function AllGirderElementsCreated() As Boolean
    AllGirderElementsCreated = Not This.girder.TopFlange Is Nothing Or _
        Not This.girder.WebPlate Is Nothing Or _
        Not This.girder.BottomFlange Is Nothing
End Function

