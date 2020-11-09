VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RolledLSectionFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@ModuleAttribute VB_Ext_KEY, "Rubberduck", "Predeclared Class Module"
'@Folder("SteelShapes.Shapes")
Option Explicit
'@PredeclaredId

Public Function Create(ByVal shapeGetter As IRolledShapeGetter, ByVal shapeName As String) As IRolledLCrossSection
    Dim newShape As RolledLCrossSection
    Set newShape = New RolledLCrossSection
    newShape.Create shapeGetter, shapeName
    Set Create = newShape
End Function
