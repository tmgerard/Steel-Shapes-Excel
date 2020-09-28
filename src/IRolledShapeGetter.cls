VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IRolledShapeGetter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Interface
'@Folder("SteelShapes.Data")
Option Explicit

'@Description "Get rolled shape data from persistant storage"
Public Function GetRolledShape(ByVal shapeName As String) As Object
Attribute GetRolledShape.VB_Description = "Get rolled shape data from persistant storage"
End Function
