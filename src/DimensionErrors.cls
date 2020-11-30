VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DimensionErrors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "Rubberduck" ,"Predeclared Class Module"
'@Folder("SteelShapes.Errors")
Option Explicit
'@PredeclaredId

Public Enum DimensionError
    BadDimension = vbObjectError + 200
End Enum

'@Description "Error due to an invalid dimension (e.g. zero or negative)"
Public Sub ThrowBadDimensionError(ByRef errorMessage As String, ByRef errorSource As String)
Attribute ThrowBadDimensionError.VB_Description = "Error due to an invalid dimension (e.g. zero or negative)"
    Err.Raise Number:=DimensionError.BadDimension, _
              Source:=errorSource, _
              Description:=errorMessage
End Sub
