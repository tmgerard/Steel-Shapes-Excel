VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BuiltUpIGirderErrors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("SteelShapes.Errors")
Option Explicit
'@PredeclaredId

Public Enum BuiltUpIGirderError
    BadPlateOrientation = vbObjectError + 300
    MissingPlateObject
End Enum

'@Description "Error due to an invalid plate girder component orientation"
Public Sub ThrowBadPlateOrientationError(ByRef errorMessage As String, ByRef errorSource As String)
Attribute ThrowBadPlateOrientationError.VB_Description = "Error due to an invalid plate girder component orientation"
    Err.Raise Number:=BuiltUpIGirderError.BadPlateOrientation, _
              Source:=errorSource, _
              Description:=errorMessage
End Sub

'@Description "Error thrown by builder when all plate girder elements are not defined"
Public Sub ThrowMissingPlateError(ByRef errorMessage As String, ByRef errorSource As String)
Attribute ThrowMissingPlateError.VB_Description = "Error thrown by builder when all plate girder elements are not defined"
    Err.Raise Number:=BuiltUpIGirderError.MissingPlateObject, _
              Source:=errorSource, _
              Description:=errorMessage
End Sub
